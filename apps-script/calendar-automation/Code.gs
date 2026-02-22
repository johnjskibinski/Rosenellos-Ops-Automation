/**************** CONFIG — DO NOT CHANGE ****************/

const CALENDAR_ID =
  "67c4ee19be733f603da0dbc8948e8731da244fe98a0855eb792b81748c8ded41@group.calendar.google.com";

const TEMPLATE_FILE_ID = "1WfdoSeTwr-nt-8OF6eBRs72YZ8kjQzEDf00xkmzH2Ws";

const CELL_NAME = "C2";
const CELL_ADDRESS = "C3";
const CELL_PHONE = "C4";
const CELL_DATETIME = "G2";

/**************** AUTOMATION SETTINGS ****************/

// Scan window
const PAST_DAYS = 3;
const FUTURE_DAYS = 21;

// Prevent patch→trigger→patch loops
const SELF_PATCH_GUARD_MS = 60000; // 60s

// Prevent rapid-fire trigger bursts
const TRIGGER_DEBOUNCE_MS = 15000; // 15s
const PROP_LAST_TRIGGER_MS = "SYNC_LAST_TRIGGER_MS";

// Prevent rapid-fire executions even if trigger debounce fails (belt + suspenders)
const GLOBAL_DEBOUNCE_MS = 15000; // 15s
const PROP_LAST_RUN_MS = "SYNC_MEASURES_LAST_RUN_MS";

// Color IDs (Google Calendar standard)
const MEASURE_COLOR_ID = "5";   // Banana
const SERVICE_COLOR_ID = "7";   // Peacock (allowed for CC-only lane)
const LOCK_WAIT_MS = 20000;

// ✅ DRY RUN: when true, nothing is written/created/updated anywhere.
const DRY_RUN = false;

// Measures: Hygiene runs on the full scan window.
// Measures: Packet + CompanyCam create/update runs ONLY for today-forward.
const PACKET_CC_TODAY_FORWARD_ONLY = true;

// ✅ Enable CompanyCam-only lane for installs/services/reminder-like entries (filtered)
const ENABLE_CC_ONLY_LANE_FOR_NON_MEASURE = true;

// Overrides (put these anywhere in the description)
const CC_FORCE_TAG = "CC:FORCE";
const CC_SKIP_TAG = "CC:SKIP";

/**************** COMPANYCAM SETTINGS ****************/

const CC_API_BASE = "https://api.companycam.com/v2";
const CC_WEB_PROJECT_BASE = "https://app.companycam.com/projects/";
const CC_PROP_PREFIX = "CC_FOR_EVENT_"; // legacy fallback only

/**************** LEGACY FALLBACK (optional) ****************/

const PROP_PREFIX = "PACKET_FOR_EVENT_"; // legacy fallback only

/**************** DESCRIPTION BLOCK REGEX ****************/

// NOTE: Don't use /g with .test() repeatedly (lastIndex pitfalls).
// We'll keep separate regexes for testing (no g) and replacement (with g).

const PACKET_BLOCK_RE_TEST = /📋\s*Measure Packet:\s*\n<https?:\/\/[^>]+>\s*/mi;
const CC_BLOCK_RE_TEST = /📸\s*CompanyCam:\s*\n<https?:\/\/[^>]+>\s*/mi;

const PACKET_BLOCK_RE_REPLACE = /📋\s*Measure Packet:\s*\n<https?:\/\/[^>]+>\s*/gmi;
const CC_BLOCK_RE_REPLACE = /📸\s*CompanyCam:\s*\n<https?:\/\/[^>]+>\s*/gmi;

/**************** DATE HELPER ****************/

function getTodayStart_() {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  return d;
}

/**************** INSTALL TRIGGER (RUN ONCE ONLY) ****************/

function installInstantCalendarTrigger() {
  ScriptApp.getProjectTriggers().forEach((t) => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("onCalendarEventUpdated")
    .forUserCalendar(CALENDAR_ID)
    .onEventUpdated()
    .create();
}

function onCalendarEventUpdated(e) {
  // Debounce at trigger entry to stop Calendar burst updates
  const props = PropertiesService.getScriptProperties();
  const nowMs = Date.now();
  const last = Number(props.getProperty(PROP_LAST_TRIGGER_MS) || "0");

  if (nowMs - last < TRIGGER_DEBOUNCE_MS) return;

  props.setProperty(PROP_LAST_TRIGGER_MS, String(nowMs));
  syncMeasures();
}

/**************** MAIN AUTOMATION ****************/

function syncMeasures() {
  return safeExecute_("syncMeasures", function() {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_WAIT_MS);

  try {
    // Global debounce (extra protection)
    const props = PropertiesService.getScriptProperties();
    const nowMs = Date.now();
    const lastRunMs = Number(props.getProperty(PROP_LAST_RUN_MS) || "0");
    if (nowMs - lastRunMs < GLOBAL_DEBOUNCE_MS) return;
    props.setProperty(PROP_LAST_RUN_MS, String(nowMs));

    const cal = CalendarApp.getCalendarById(CALENDAR_ID);

    const now = new Date();
    const start = new Date(now.getTime() - PAST_DAYS * 24 * 60 * 60 * 1000);
    const end = new Date(now.getTime() + FUTURE_DAYS * 24 * 60 * 60 * 1000);

    const events = cal.getEvents(start, end);
    const todayStart = getTodayStart_();

    for (const ev of events) {
      const titleRaw = (ev.getTitle() || "").trim();
      const descRaw = (ev.getDescription() || "").trim();
      const locationRaw = (ev.getLocation() || "").trim();

      const eventIdFull = ev.getId();
      const eventIdBase = eventIdFull.split("@")[0];

      // ✅ Calendar API: ONE GET per event
      const evRes = Calendar.Events.get(CALENDAR_ID, eventIdBase);
      const priv = evRes.extendedProperties?.private || {};

      // Per-event loop guard: skip events we just patched
      const lastPatchedBy = String(priv.LAST_PATCH_BY || "");
      const lastPatchedAt = String(priv.LAST_PATCH_AT || "");
      if (lastPatchedBy && lastPatchedAt) {
        const age = Date.now() - new Date(lastPatchedAt).getTime();
        if (age >= 0 && age < SELF_PATCH_GUARD_MS) {
          continue;
        }
      }

      const color = String(ev.getColor());

      /**********************************************************************
       * LANE A: MEASURES (Banana) — keep your existing behavior
       **********************************************************************/
      if (color === MEASURE_COLOR_ID) {
        const startTime = ev.getStartTime();
        const isTodayForward = startTime >= todayStart;
        const allowPacketCc = !PACKET_CC_TODAY_FORWARD_ONLY || isTodayForward;

        const privateUpdates = { ...priv };
        const actions = [];

        /********** DATA HYGIENE (ALWAYS IN SCAN WINDOW) **********/
        const hygiene = applyEventDataHygiene_({
          title: titleRaw,
          desc: descRaw,
          location: locationRaw,
        });

        const title = hygiene.title;
        const descHygiene = hygiene.desc;
        const location = hygiene.location;

        if (title !== titleRaw) actions.push("patched_title");
        if (location !== locationRaw) actions.push("patched_location");
        if (descHygiene !== descRaw) actions.push("patched_desc_hygiene");

        // Canonical display formats
        const nameFormats = extractNameFormats_(title);
        const customerName = (nameFormats.firstLast || title || "Project").trim(); // First Last (CC preference)
        const displayName = (nameFormats.lastFirst || title || "Project").trim(); // Last, First

        // Canonicalize address for packet + CC (reduces drift over time)
        const canonicalAddress = canonicalizeAddressOneLine_(location);

        // Phone from description (hygiene moved title phone here)
        const phone = parsePhone_(descHygiene);

        /********** PACKET & COMPANYCAM (TODAY FORWARD ONLY) **********/
        const existingPacketId =
          (privateUpdates.PACKET_FILE_ID ? String(privateUpdates.PACKET_FILE_ID) : "") ||
          props.getProperty(PROP_PREFIX + eventIdFull) ||
          "";

        const hasPacketBlock = PACKET_BLOCK_RE_TEST.test(descHygiene);

        let fileIdToUse = existingPacketId;
        let packetWasCreated = false;

        if (allowPacketCc) {
          if (!fileIdToUse) {
            if (DRY_RUN) {
              log_(
                `[DRY_RUN] Would CREATE packet for "${customerName}" | phone="${phone}" | addr="${canonicalAddress}"`
              );
            } else {
              const packet = createPacketFromTemplate_({
                customerName,
                displayName,
                phone,
                address: canonicalAddress,
                datetime: startTime,
              });
              fileIdToUse = packet.fileId;
              packetWasCreated = true;

              props.setProperty(PROP_PREFIX + eventIdFull, fileIdToUse);
              privateUpdates.PACKET_FILE_ID = String(fileIdToUse);
              actions.push("created_packet");
            }
          } else {
            if (DRY_RUN) {
              log_(
                `[DRY_RUN] Would UPDATE packet ${fileIdToUse} for "${customerName}" | phone="${phone}" | addr="${canonicalAddress}"`
              );
            } else {
              updateExistingPacket_(fileIdToUse, {
                customerName,
                phone,
                address: canonicalAddress,
                datetime: startTime,
              });
              if (!props.getProperty(PROP_PREFIX + eventIdFull)) {
                props.setProperty(PROP_PREFIX + eventIdFull, fileIdToUse);
              }
              privateUpdates.PACKET_FILE_ID = String(fileIdToUse);
              actions.push("updated_packet");
            }
          }
        }

        // Build desired system blocks only when we’re allowed to do packet/cc work
        let desiredPacketBlock = "";
        if (allowPacketCc && fileIdToUse) {
          const packetUrl = buildDriveOpenUrl_(fileIdToUse);
          desiredPacketBlock = buildPacketBlock_(packetUrl);
        }

        const hasCcBlock = CC_BLOCK_RE_TEST.test(descHygiene);
        const shouldDoCcWork =
          allowPacketCc &&
          !!canonicalAddress &&
          (packetWasCreated || !hasPacketBlock || !hasCcBlock);

        let desiredCcBlock = "";
        let ccProjectId = privateUpdates.CC_PROJECT_ID ? String(privateUpdates.CC_PROJECT_ID) : "";

        if (shouldDoCcWork) {
          const cc = ensureCompanyCamProjectDuplicateProofForever_({
            props,
            eventIdFull,
            privateUpdates,
            customerName,
            calendarAddress: canonicalAddress,
          });

          if (cc?.projectId) {
            ccProjectId = String(cc.projectId);
            privateUpdates.CC_PROJECT_ID = ccProjectId;
            privateUpdates.CC_MATCH_KEY = cc.matchKey || "";
            privateUpdates.CC_LAST_SYNC = new Date().toISOString();

            let ccUrl = CC_WEB_PROJECT_BASE + ccProjectId;
            if (!ccUrl.endsWith("/")) ccUrl += "/";
            desiredCcBlock = buildCompanyCamBlock_(ccUrl);
            actions.push("linked_companycam");
          } else {
            privateUpdates.CC_NEEDS_REVIEW = "1";
            privateUpdates.CC_REVIEW_REASON = cc?.reason || "NO_SAFE_MATCH_AND_NOT_CREATED";
            actions.push("companycam_needs_review");
          }
        } else if (allowPacketCc && ccProjectId) {
          let ccUrl = CC_WEB_PROJECT_BASE + ccProjectId;
          if (!ccUrl.endsWith("/")) ccUrl += "/";
          desiredCcBlock = buildCompanyCamBlock_(ccUrl);
        }

        /********** DESCRIPTION FINALIZATION (SYSTEM BLOCKS) **********/
        let finalDesc = descHygiene;

        if (allowPacketCc) {
          let cleaned = removeSystemBlocks_(finalDesc);

          const blocks = [];
          if (desiredPacketBlock) blocks.push(desiredPacketBlock);
          if (desiredCcBlock) blocks.push(desiredCcBlock);

          finalDesc = blocks.length ? joinWithSpacing_(cleaned, blocks.join("\n")) : cleaned;
          if (finalDesc !== descHygiene) actions.push("rebuilt_system_blocks");
        }

        /********** PATCH CALENDAR EVENT (ONE PATCH MAX) **********/
        const patch = {};

        if (title !== titleRaw) patch.summary = title;

        // Patch to canonical formatting (compare canonical-to-canonical to avoid drift)
        const locCanonRaw = canonicalizeAddressOneLine_(locationRaw);
        if (canonicalAddress !== locCanonRaw) patch.location = canonicalAddress;

        if (finalDesc !== descRaw) patch.description = finalDesc;

        // Only stamp audit fields if we are already changing visible fields (prevents metadata-only patch loops)
        const isChangingVisibleFields = !!(patch.summary || patch.location || patch.description);
        if (actions.length && isChangingVisibleFields) {
          privateUpdates.LAST_AUTOMATION_RUN = new Date().toISOString();
          privateUpdates.LAST_AUTOMATION_ACTIONS = actions.join(",");
        }

        if (Object.keys(patch).length) {
          privateUpdates.LAST_PATCH_BY = "syncMeasures";
          privateUpdates.LAST_PATCH_AT = new Date().toISOString();
          patch.extendedProperties = { private: privateUpdates };

          if (DRY_RUN) {
            log_(`[DRY_RUN] Would PATCH MEASURE event ${eventIdBase}: ${JSON.stringify(patch)}`);
          } else {
            Calendar.Events.patch(patch, CALENDAR_ID, eventIdBase);
          }
        }

        continue; // ✅ done with this event
      }

      /**********************************************************************
       * LANE B: NON-MEASURE (Installs/Services) — CompanyCam ONLY
       * - No title changes
       * - No location changes
       * - No packet
       * - Append CC link block at bottom
       **********************************************************************/
      if (!ENABLE_CC_ONLY_LANE_FOR_NON_MEASURE) continue;

      const privateUpdates = { ...priv };

      // Overrides
      const override = getCcOverride_(descRaw);
      if (override === "SKIP") continue;

      const isServiceColor = (color === SERVICE_COLOR_ID);
      const looksLikeJob = isInstallOrServiceTitle_(titleRaw, descRaw);
      const isEligible = (override === "FORCE") || isServiceColor || looksLikeJob;

      if (!isEligible) continue;

      // Skip obvious reminders / non-job items unless FORCE
      if (override !== "FORCE" && isReminderLike_(titleRaw, descRaw)) continue;

      // Already has CC block? do nothing
      if (CC_BLOCK_RE_TEST.test(descRaw)) continue;

      // Need an address to do safe duplicate-proof matching/creation
      const addrRaw = pickBestAddressForNonMeasure_(locationRaw, titleRaw, descRaw);

      // ✅ FIX: normalize one-line Google locations into "street, city, state zip"
      const normalizedAddr = normalizeLocationToOneLine_(addrRaw);
      const canonicalAddress = canonicalizeAddressOneLine_(normalizedAddr);

      if (!canonicalAddress || !addressHasStreetAndCity_(canonicalAddress)) {
        // Without street+city we avoid creating or matching (too risky for duplicates)
        continue;
      }

      // Customer name for CompanyCam (Firstname Lastname)
      const customerName = buildCustomerNameForNonMeasure_(titleRaw);

      const cc = ensureCompanyCamProjectDuplicateProofForever_({
        props,
        eventIdFull,
        privateUpdates,
        customerName,
        calendarAddress: canonicalAddress,
      });

      const ccProjectId = cc?.projectId ? String(cc.projectId) : "";
      if (!ccProjectId) continue;

      // Build CC link
      let ccUrl = CC_WEB_PROJECT_BASE + ccProjectId;
      if (!ccUrl.endsWith("/")) ccUrl += "/";
      const ccBlock = buildCompanyCamBlock_(ccUrl);

      // Append at very bottom, leave existing content untouched
      const finalDesc = appendBlockToBottom_(descRaw, ccBlock);

      const patch = {};
      patch.description = finalDesc;

      // Stamp/keep metadata
      privateUpdates.CC_PROJECT_ID = ccProjectId;
      privateUpdates.CC_MATCH_KEY = cc?.matchKey || privateUpdates.CC_MATCH_KEY || "";
      privateUpdates.CC_LAST_SYNC = new Date().toISOString();

      privateUpdates.LAST_PATCH_BY = "syncMeasures";
      privateUpdates.LAST_PATCH_AT = new Date().toISOString();
      patch.extendedProperties = { private: privateUpdates };

      if (DRY_RUN) {
        log_(`[DRY_RUN] Would PATCH NON-MEASURE event ${eventIdBase}: ${JSON.stringify(patch)}`);
      } else {
        Calendar.Events.patch(patch, CALENDAR_ID, eventIdBase);
      }
    }
  } finally {
    lock.releaseLock();
  }
  });
}

/**************** LANE B HELPERS (NON-MEASURE CC ONLY) ****************/

function getCcOverride_(desc) {
  const d = String(desc || "").toUpperCase();
  if (d.indexOf(CC_SKIP_TAG.toUpperCase()) !== -1) return "SKIP";
  if (d.indexOf(CC_FORCE_TAG.toUpperCase()) !== -1) return "FORCE";
  return "";
}

function isReminderLike_(title, desc) {
  const t = normalizeLoose_(title);
  const d = normalizeLoose_(desc);
  const hay = `${t}\n${d}`;

  // Things you explicitly wanted skipped (unless forced)
  const badPhrases = [
    "PAINT SAMPLE",
    "PAINT SAMPLES",
    "SAMPLES",
    "PICK UP PARTS",
    "PICKUP PARTS",
    "PICK-UP PARTS",
    "PICK UP",
    "PICKUP",
    "CALL ",
    "CALL-",
    "TEXT ",
    "EMAIL ",
    "PAYMENT",
    "INVOICE",
    "REMINDER",
    "FOLLOW UP",
    "FOLLOW-UP",
    "MEETING",
    "ESTIMATE",
    "QUOTE",
  ];

  for (const p of badPhrases) {
    if (hay.indexOf(p) !== -1) return true;
  }
  return false;
}

function isInstallOrServiceTitle_(title, desc) {
  const t = normalizeLoose_(title);
  const d = normalizeLoose_(desc);
  const hay = `${t}\n${d}`;

  // Strong shorthand signals (catches "1DHs" and "20 DHs" etc)
  const tokenRe = /\b\d+\s*(DH|DHS|SLIDER|SLIDERS|CASEMENT|CASEMENTS|PIC|PICTURE|TWIN|TRIPLE|STORM\s*DOOR|STORM\s*DOORS|DOOR|DOORS|WINDOW|WINDOWS)\b/;
  if (tokenRe.test(hay)) return true;

  // Positive job signals (your install titles strongly match these)
  const good = [
    "DH",
    "DHS",
    "SLIDER",
    "SLIDERS",
    "CASEMENT",
    "CASEMENTS",
    "BAY",
    "BOX BAY",
    "BOW",
    "PIC",
    "PICTURE",
    "DOOR",
    "STORM DOOR",
    "CAPPING",
    "CAP ",
    "INSTALL",
    "REMOVE AND INSTALL",
    "FULL FRAME",
    "WINDOW",
    "WINDOWS",
    "SIDING",
    "ROOF",
    "GUTTER",
    "SOFFIT",
    "FASCIA",
    "TRIM",
    "SERVICE",
    "REPAIR",
  ];

  for (const g of good) {
    if (hay.indexOf(g) !== -1) return true;
  }

  // Dash structure + counts + window terms
  if (
    hay.indexOf(" - ") !== -1 &&
    /\b\d+\b/.test(hay) &&
    (hay.indexOf("DH") !== -1 || hay.indexOf("SLIDER") !== -1 || hay.indexOf("WINDOW") !== -1 || hay.indexOf("DOOR") !== -1)
  ) {
    return true;
  }

  return false;
}

function pickBestAddressForNonMeasure_(locationRaw, titleRaw, descRaw) {
  const loc = String(locationRaw || "").trim();
  if (loc) return loc;

  // Fallback: sometimes titles/descriptions contain an address
  const combo = `${titleRaw}\n${descRaw}`.trim();
  const found = extractAddressFromAnyText_(combo);
  return found?.address || "";
}

/**
 * ✅ FIX: Convert Google Maps-ish one-liners like:
 *   "113 Main St Bridgeport, NJ 08014"
 * into:
 *   "113 Main St, Bridgeport, NJ 08014"
 * Also strips trailing ", USA".
 */
function normalizeLocationToOneLine_(raw) {
  let x = String(raw || "").trim();
  if (!x) return "";

  x = x.replace(/\s*,\s*USA\s*$/i, "").trim();

  // If we can parse an address from it, return the parsed one-line version
  const fromComma = extractAddressCommaBased_(x);
  if (fromComma?.address) return fromComma.address;

  const fromSpace = extractAddressSpaceBased_(x);
  if (fromSpace?.address) return fromSpace.address;

  // If it's already comma-ish, keep it
  return x;
}

function addressHasStreetAndCity_(addressOneLine) {
  const normalized = normalizeLocationToOneLine_(addressOneLine);
  if (!normalized) return false;

  const parsed = parseUsAddressToCompanyCam_(normalized);
  const street = String(parsed.street_address_1 || "").trim();
  const city = String(parsed.city || "").trim();

  return !!street && !!city;
}

function buildCustomerNameForNonMeasure_(titleRaw) {
  // Remove leading crew tags like "(JW,MJ)" "(Matt & MJ)" "(IN HOUSE)" "(Chuck)" "(Push?)"
  let t = String(titleRaw || "").trim();

  t = t.replace(/^\s*\(([^)]+)\)\s*/g, "").trim();

  // Take left side before first " - "
  const left = t.split(" - ")[0].trim();

  // Remove ETA parentheticals, etc.
  const cleaned = left.replace(/\([^)]*\)/g, " ").replace(/\s+/g, " ").trim();

  // Convert "Last, First" to "First Last" (CompanyCam preference)
  const formats = extractNameFormats_(cleaned);
  return (formats.firstLast || cleaned || "Project").trim();
}

function appendBlockToBottom_(desc, block) {
  const d = String(desc || "").trim();
  const b = String(block || "").trim();
  if (!b) return d;

  if (!d) return b;
  return (d + "\n\n" + b).trim();
}

function normalizeLoose_(s) {
  return String(s || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

/**************** DATA HYGIENE HELPERS (MEASURES) ****************/

function applyEventDataHygiene_({ title, desc, location }) {
  const origTitle = String(title || "").trim();
  const origDesc = String(desc || "").trim();
  const origLoc = String(location || "").trim();

  let t = origTitle;

  // 1) Remove "measure" variants from title (delete, never move)
  t = stripMeasureWord_(t);

  // 2) Extract phone from title
  const phoneFromTitle = extractPhoneFromText_(t);
  if (phoneFromTitle) t = removeSubstringBestEffort_(t, phoneFromTitle.raw);

  // 3) Extract FULL address from title (street + city required)
  const addrObj = extractAddressFromTitle_(t);
  if (addrObj && addrObj.rawMatch) t = removeSubstringBestEffort_(t, addrObj.rawMatch);

  // 4) Name chunk = left-most segment before separators
  const nameChunk = t.split(/[-–—|]/)[0].trim();
  const cleanTitle = cleanNameTitle_(nameChunk) || nameChunk || origTitle || "Project";

  // 5) Remaining title content becomes notes
  const extraChunk = t.replace(nameChunk, "").replace(/^[\s\-\–—|:]+/, "").trim();

  // 6) New location + desc
  let newLoc = origLoc;
  let newDesc = origDesc;

  if (addrObj && addrObj.address) {
    const extractedAddr = canonicalizeAddressOneLine_(addrObj.address);
    if (!newLoc) {
      newLoc = extractedAddr;
    } else {
      const same = addressesRoughlySame_(newLoc, extractedAddr);
      if (!same) {
        newDesc = appendNoteLine_(newDesc, `Title contained address: ${extractedAddr}`);
      }
    }
  }

  // 7) Phone -> description
  if (phoneFromTitle && phoneFromTitle.normalized) {
    newDesc = upsertPhoneLine_(newDesc, phoneFromTitle.normalized);
  }

  // 8) Extra title content -> description
  if (extraChunk) {
    newDesc = appendNoteLine_(newDesc, `Title notes: ${extraChunk}`);
  }

  return {
    title: cleanTitle.trim(),
    desc: newDesc.trim(),
    location: newLoc.trim(),
  };
}

function stripMeasureWord_(s) {
  let x = String(s || "");
  x = x.replace(/\b[-–—]?\s*\(?\s*measure\s*(appt|appointment)?\s*\)?\b/gi, " ");
  x = x.replace(/\s{2,}/g, " ").trim();
  return x;
}

function extractPhoneFromText_(s) {
  const raw = String(s || "");
  const m = raw.match(/(\+?1?[\s\-\.]?\(?\d{3}\)?[\s\-\.]?\d{3}[\s\-\.]?\d{4})/);
  if (!m) return null;
  const found = m[1].trim();
  return { raw: found, normalized: normalizePhone_(found) };
}

function normalizePhone_(s) {
  const digits = String(s || "").replace(/\D/g, "");
  const ten = bestUs10DigitPhone_(digits);
  if (!ten) return "";
  return `(${ten.slice(0, 3)}) ${ten.slice(3, 6)}-${ten.slice(6)}`;
}

function bestUs10DigitPhone_(digitsOnly) {
  let d = String(digitsOnly || "").replace(/\D/g, "");
  if (!d) return "";

  if (d.length === 10) return d;
  if (d.length === 11 && d.startsWith("1")) return d.slice(1);

  if (d.length > 11) {
    if (d.startsWith("1") && d.length >= 11) {
      const cand = d.slice(1, 11);
      if (isPlausibleUsPhone10_(cand)) return cand;
    }
    for (let i = 0; i <= d.length - 10; i++) {
      const cand = d.slice(i, i + 10);
      if (isPlausibleUsPhone10_(cand)) return cand;
    }
  }

  if (d.length === 11) {
    for (let i = 0; i <= 1; i++) {
      const cand = d.slice(i, i + 10);
      if (isPlausibleUsPhone10_(cand)) return cand;
    }
  }

  return "";
}

function isPlausibleUsPhone10_(tenDigits) {
  if (!/^\d{10}$/.test(tenDigits)) return false;
  const area = tenDigits.slice(0, 3);
  const exch = tenDigits.slice(3, 6);

  if (area[0] === "0" || area[0] === "1") return false;
  if (exch[0] === "0" || exch[0] === "1") return false;

  return true;
}

function extractAddressFromTitle_(t) {
  const raw = String(t || "").trim();
  if (!raw) return null;

  const commaMatch = extractAddressCommaBased_(raw);
  if (commaMatch) return commaMatch;

  const spaceMatch = extractAddressSpaceBased_(raw);
  if (spaceMatch) return spaceMatch;

  return null;
}

// Used by NON-MEASURE lane too (general text)
function extractAddressFromAnyText_(text) {
  const raw = String(text || "").trim();
  if (!raw) return null;

  // Try line-by-line first
  const lines = raw.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  for (const line of lines) {
    const commaMatch = extractAddressCommaBased_(line);
    if (commaMatch) return commaMatch;
    const spaceMatch = extractAddressSpaceBased_(line);
    if (spaceMatch) return spaceMatch;
  }

  // Then whole blob
  const commaMatch = extractAddressCommaBased_(raw);
  if (commaMatch) return commaMatch;
  const spaceMatch = extractAddressSpaceBased_(raw);
  if (spaceMatch) return spaceMatch;

  return null;
}

function extractAddressCommaBased_(raw) {
  const parts = raw.split(",").map((p) => p.trim()).filter(Boolean);
  if (parts.length < 2) return null;

  for (let i = 0; i < parts.length - 1; i++) {
    const street = extractStreetLine_(parts[i]);
    if (!street) continue;

    const cityStateZip = parts[i + 1];
    const parsedCSZ = parseCityStateZip_(cityStateZip);

    let city = parsedCSZ.city;
    let state = parsedCSZ.state;
    let zip = parsedCSZ.zip;

    if (parts.length >= i + 3 && (!state || !zip)) {
      const parsed2 = parseCityStateZip_(parts[i + 2]);
      if (!state && parsed2.state) state = parsed2.state;
      if (!zip && parsed2.zip) zip = parsed2.zip;
    }

    if (!city) continue;

    const address = formatOneLineAddress_(street, city, state, zip);
    const rawMatch = [parts[i], parts[i + 1], parts[i + 2]].filter(Boolean).join(", ");
    return { address, street, city, state, zip, rawMatch };
  }
  return null;
}

function extractAddressSpaceBased_(raw) {
  const street = extractStreetLine_(raw);
  if (!street) return null;

  const idx = raw.toLowerCase().indexOf(street.toLowerCase());
  if (idx === -1) return null;

  let tail = raw.slice(idx + street.length).trim();
  if (!tail) return null;

  tail = tail.split(/[-–—|]/)[0].trim();
  tail = tail
    .replace(/(\+?1?[\s\-\.]?\(?\d{3}\)?[\s\-\.]?\d{3}[\s\-\.]?\d{4})/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  // If tail starts with a state/country marker, reject
  const parsed = parseCityStateZip_(tail);
  if (!parsed.city) return null;

  const address = formatOneLineAddress_(street, parsed.city, parsed.state, parsed.zip);
  const rawMatch = `${street} ${tail}`.trim();
  return { address, street, city: parsed.city, state: parsed.state, zip: parsed.zip, rawMatch };
}

function extractStreetLine_(s) {
  const raw = String(s || "").trim();
  if (!raw) return "";
  const re =
    /\b(\d{1,6})\s+([A-Z0-9.\-']+\s+){0,8}(STREET|ST|AVENUE|AVE|ROAD|RD|DRIVE|DR|LANE|LN|COURT|CT|PLACE|PL|TERRACE|TER|CIRCLE|CIR|BOULEVARD|BLVD|PARKWAY|PKWY|HIGHWAY|HWY)\b\.?/i;
  const m = raw.match(re);
  return m ? m[0].trim().replace(/\.$/, "") : "";
}

function parseCityStateZip_(s) {
  let x = String(s || "").trim();
  if (!x) return { city: "", state: "", zip: "" };

  x = x.replace(/,/g, " ").replace(/\s+/g, " ").trim();

  let zip = "";
  const zipM = x.match(/\b(\d{5}(?:-\d{4})?)\b$/);
  if (zipM) {
    zip = zipM[1];
    x = x.slice(0, zipM.index).trim();
  }

  let state = "";
  const stM = x.match(/\b([A-Z]{2})\b$/i);
  if (stM) {
    state = stM[1].toUpperCase();
    x = x.slice(0, stM.index).trim();
  }

  const city = sanitizeCity_(x);
  return { city, state, zip };
}

function sanitizeCity_(s) {
  const x = String(s || "").trim();
  if (!x) return "";

  const cleaned = x
    .replace(/\b(measure|appt|appointment)\b/gi, "")
    .replace(/\b(usa)\b/gi, "")
    .replace(/\s+/g, " ")
    .trim();
  if (!cleaned) return "";

  const parts = cleaned.split(" ").filter(Boolean).slice(0, 4);
  const candidate = parts.join(" ");
  if (!/^[A-Z.'\- ]+$/i.test(candidate)) return "";
  return candidate;
}

function formatOneLineAddress_(street, city, state, zip) {
  const st = String(street || "").trim();
  const c = String(city || "").trim();
  const s = String(state || "").trim();
  const z = String(zip || "").trim();
  const csz = [c, [s, z].filter(Boolean).join(" ")].filter(Boolean).join(", ");
  return [st, csz].filter(Boolean).join(", ").trim();
}

function canonicalizeAddressOneLine_(address) {
  const raw = String(address || "").trim();
  if (!raw) return "";
  const obj = parseUsAddressToCompanyCam_(raw);
  const one = formatAddressObjOneLine_(obj);
  return one && one.length >= 6 ? one.trim() : raw;
}

function addressesRoughlySame_(a, b) {
  const aa = canonicalizeAddressOneLine_(a);
  const bb = canonicalizeAddressOneLine_(b);

  if (normalizeAddress_(aa) === normalizeAddress_(bb)) return true;

  const aKey = extractStreetKey_(aa);
  const bKey = extractStreetKey_(bb);
  if (aKey && bKey && normalizeAddress_(aKey) === normalizeAddress_(bKey)) {
    const aZip = (aa.match(/\b\d{5}(?:-\d{4})?\b/) || [])[0] || "";
    const bZip = (bb.match(/\b\d{5}(?:-\d{4})?\b/) || [])[0] || "";
    if (aZip && bZip) return aZip === bZip;
    return true;
  }

  return false;
}

function cleanNameTitle_(nameChunk) {
  const x = String(nameChunk || "").trim();
  if (!x) return "";
  if (x.includes(",")) return x.replace(/\s{2,}/g, " ").trim();
  return x.replace(/\s{2,}/g, " ").trim();
}

function removeSubstringBestEffort_(whole, part) {
  if (!whole || !part) return String(whole || "");
  let out = String(whole);
  const idx = out.toLowerCase().indexOf(String(part).toLowerCase());
  if (idx !== -1) {
    out = (out.slice(0, idx) + " " + out.slice(idx + part.length))
      .replace(/\s{2,}/g, " ")
      .trim();
  }
  return out;
}

function upsertPhoneLine_(desc, phone) {
  const d = String(desc || "");
  if (!phone) return d;

  if (/^phone\s*:/im.test(d)) {
    return d.replace(/^phone\s*:\s*.*$/im, `Phone: ${phone}`).trim();
  }
  return (`Phone: ${phone}\n` + d).replace(/\n{3,}/g, "\n\n").trim();
}

function appendNoteLine_(desc, line) {
  const d = String(desc || "").trim();
  const l = String(line || "").trim();
  if (!l) return d;
  if (!d) return l;
  return (d + "\n" + l).trim();
}

function log_(msg) {
  Logger.log(msg);
  console.log(msg);
}

/**************** DESCRIPTION BUILDERS ****************/

function buildPacketBlock_(url) {
  return `📋 Measure Packet:\n<${String(url || "").trim()}>`;
}

function buildCompanyCamBlock_(url) {
  return `📸 CompanyCam:\n<${String(url || "").trim()}>`;
}

function removeSystemBlocks_(desc) {
  let out = (desc || "").trim();
  if (!out) return "";

  out = out.replace(PACKET_BLOCK_RE_REPLACE, "").trim();
  out = out.replace(CC_BLOCK_RE_REPLACE, "").trim();

  // legacy one-line variants
  out = out.replace(/^📋\s*Measure Packet\s*:.*$/gmi, "").trim();
  out = out.replace(/^📸\s*CompanyCam\s*:.*$/gmi, "").trim();

  // legacy junk
  out = out.replace(/^\[PACKET:[a-zA-Z0-9_-]{10,}\]\s*$/gmi, "").trim();
  out = out.replace(/^.*OPEN\s+MEASURE\s+PACKET.*$/gmi, "").trim();
  out = out.replace(/^.*COMPANYCAM\s*\(TEAM\).*$/gmi, "").trim();

  out = out.replace(/\n{3,}/g, "\n\n").trim();
  return out;
}

function joinWithSpacing_(base, addition) {
  const a = (base || "").trim();
  const b = (addition || "").trim();
  if (!a) return b;
  if (!b) return a;
  return (a + "\n\n" + b).trim();
}

/**************** COMPANYCAM: DUPLICATE-PROOF FOREVER ****************/

function ensureCompanyCamProjectDuplicateProofForever_({
  props,
  eventIdFull,
  privateUpdates,
  customerName,
  calendarAddress,
}) {
  const existing = privateUpdates.CC_PROJECT_ID ? String(privateUpdates.CC_PROJECT_ID) : "";
  if (existing) return { projectId: existing, matchKey: privateUpdates.CC_MATCH_KEY || "" };

  const legacy = props.getProperty(CC_PROP_PREFIX + eventIdFull) || "";
  if (legacy) return { projectId: legacy, matchKey: "" };

  const addrObj = parseUsAddressToCompanyCam_(calendarAddress);

  const wantedKey = streetKeyFromAnyAddress_(addrObj, calendarAddress, { name: "" });
  const wantedKeyNorm = normalizeAddress_(wantedKey || "");
  const wantedNum = extractStreetNumber_(calendarAddress);
  const wantedStem = extractStreetStem_(calendarAddress);

  const queries = buildCompanyCamQueryPlan_(addrObj, calendarAddress, wantedKey, wantedNum, wantedStem);

  const byId = {};
  for (const q of queries) {
    const list = listCompanyCamProjectsByQuery_(q);
    for (const p of list) {
      const id = String(p.id);
      if (!byId[id]) byId[id] = p;
    }
  }

  const candidates = Object.keys(byId).map((id) => byId[id]);

  if (candidates.length) {
    const ranked = rankCompanyCamCandidates_(candidates, {
      wantedKeyNorm,
      wantedNum,
      wantedStem,
      rawAddress: calendarAddress,
    });

    const best = ranked[0];
    if (best && best.score >= 80) {
      const id = String(best.project.id);
      if (!DRY_RUN) props.setProperty(CC_PROP_PREFIX + eventIdFull, id);
      return { projectId: id, matchKey: wantedKey || best.matchedKey || "" };
    }

    return {
      projectId: "",
      reason: `CANDIDATES_FOUND_BUT_NOT_CONFIDENT (count=${candidates.length}, bestScore=${
        best ? best.score : "n/a"
      })`,
    };
  }

  if (DRY_RUN) {
    log_(
      `[DRY_RUN] Would CREATE CompanyCam project: name="${customerName}" address="${formatAddressObjOneLine_(addrObj)}"`
    );
    return { projectId: "", matchKey: wantedKey || "" };
  }

  const created = createCompanyCamProject_(customerName, addrObj);
  const id = String(created.id);
  props.setProperty(CC_PROP_PREFIX + eventIdFull, id);
  return { projectId: id, matchKey: wantedKey || "" };
}

function buildCompanyCamQueryPlan_(addrObj, rawAddress, wantedKey, wantedNum, wantedStem) {
  const out = [];
  const street1 = (addrObj.street_address_1 || "").trim();
  const raw = String(rawAddress || "").trim();

  if (street1) out.push(street1);
  if (wantedKey) out.push(wantedKey);
  if (wantedNum && wantedStem) out.push(`${wantedNum} ${wantedStem}`);
  if (raw) out.push(raw);

  const seen = {};
  return out.filter((x) => {
    const k = normalizeAddress_(x);
    if (!k) return false;
    if (seen[k]) return false;
    seen[k] = true;
    return true;
  });
}

function rankCompanyCamCandidates_(projects, ctx) {
  const ranked = [];

  for (const p of projects) {
    const addrKey = streetKeyFromAnyAddress_(p.address || null, "", p);
    const addrKeyNorm = normalizeAddress_(addrKey || "");
    const nameNorm = normalizeAddress_(p?.name || "");

    let score = 0;

    if (ctx.wantedKeyNorm && addrKeyNorm && addrKeyNorm === ctx.wantedKeyNorm) score += 100;

    const pNum = extractStreetNumber_(addrKey || p?.name || "");
    const pStem = extractStreetStem_(addrKey || p?.name || "");

    if (ctx.wantedNum && pNum && ctx.wantedNum === pNum) score += 30;
    if (ctx.wantedStem && pStem && ctx.wantedStem === pStem) score += 40;

    if (ctx.wantedKeyNorm && nameNorm && nameNorm.indexOf(ctx.wantedKeyNorm) !== -1) score += 35;
    if (
      ctx.wantedNum &&
      ctx.wantedStem &&
      nameNorm &&
      nameNorm.indexOf(normalizeAddress_(`${ctx.wantedNum} ${ctx.wantedStem}`)) !== -1
    )
      score += 25;

    if (!addrKeyNorm && !nameNorm) score -= 20;

    ranked.push({ project: p, score, matchedKey: addrKey || "" });
  }

  ranked.sort((a, b) => b.score - a.score);
  return ranked;
}

/**************** COMPANYCAM API HELPERS ****************/

function listCompanyCamProjectsByQuery_(queryStr) {
  const q = encodeURIComponent(String(queryStr || "").trim());
  if (!q) return [];
  const resp = ccFetch_(`/projects?query=${q}&per_page=100&page=1`, "get");
  return normalizeProjectList_(resp);
}

function createCompanyCamProject_(name, addressObj) {
  return ccFetch_("/projects", "post", {
    name: String(name || "Project").trim(),
    address: addressObj,
  });
}

function ccFetch_(path, method, bodyObj) {
  const token = PropertiesService.getScriptProperties().getProperty("COMPANYCAM_TOKEN");
  if (!token) throw new Error("Missing Script Property: COMPANYCAM_TOKEN");

  const url = CC_API_BASE + path;

  const opts = {
    method,
    muteHttpExceptions: true,
    headers: { Authorization: "Bearer " + token, accept: "application/json" },
  };
  if (bodyObj) {
    opts.headers["content-type"] = "application/json";
    opts.payload = JSON.stringify(bodyObj);
  }

  const maxAttempts = 5;
  let attempt = 0;

  while (true) {
    attempt++;
    const resp = UrlFetchApp.fetch(url, opts);
    const code = resp.getResponseCode();
    const text = resp.getContentText();

    if (code >= 200 && code < 300) {
      return text ? JSON.parse(text) : null;
    }

    const isRetryable = code === 429 || (code >= 500 && code <= 599);
    if (!isRetryable || attempt >= maxAttempts) {
      throw new Error(`CompanyCam API error ${code}: ${text}`);
    }

    const baseMs = 500 * Math.pow(2, attempt - 1);
    const jitter = Math.floor(Math.random() * 250);
    Utilities.sleep(baseMs + jitter);
  }
}

function normalizeProjectList_(projectsResponse) {
  if (!projectsResponse) return [];
  if (Array.isArray(projectsResponse)) return projectsResponse;
  if (Array.isArray(projectsResponse.projects)) return projectsResponse.projects;
  if (Array.isArray(projectsResponse.data)) return projectsResponse.data;
  return [];
}

/**************** ADDRESS PARSING + STREET KEY ****************/

function parseUsAddressToCompanyCam_(addressRaw) {
  const raw = String(addressRaw || "").trim();
  if (!raw) {
    return {
      street_address_1: "",
      street_address_2: "",
      city: "",
      state: "",
      postal_code: "",
      country: "US",
    };
  }

  // Remove trailing ", USA" if present (helps consistency)
  const cleaned = raw.replace(/\s*,\s*USA\s*$/i, "").trim();
  const parts = cleaned.split(",").map((p) => p.trim()).filter(Boolean);

  let street1 = parts[0] || cleaned;
  let city = parts[1] || "";
  let state = "";
  let postal = "";

  if (parts.length >= 3) {
    const stateZip = parts[2];
    const m = stateZip.match(/^([A-Z]{2})\s*(\d{5}(?:-\d{4})?)?/i);
    if (m) {
      state = (m[1] || "").toUpperCase();
      postal = m[2] || "";
    }
  }

  if (!state && parts.length >= 2) {
    const m = parts[1].match(/^(.+?)\s+([A-Z]{2})\s*(\d{5}(?:-\d{4})?)?/i);
    if (m) {
      city = (m[1] || "").trim();
      state = (m[2] || "").toUpperCase();
      postal = m[3] || "";
    }
  }

  return {
    street_address_1: street1,
    street_address_2: "",
    city,
    state,
    postal_code: postal,
    country: "US",
  };
}

function streetKeyFromAnyAddress_(addrObjOrString, fallbackString, projectObj) {
  if (addrObjOrString && typeof addrObjOrString === "object") {
    const s1 = (addrObjOrString.street_address_1 || "").trim();
    if (s1) {
      const k = extractStreetKey_(s1);
      if (k) return k;
    }
    const formatted = formatAddressObjOneLine_(addrObjOrString);
    const k2 = extractStreetKey_(formatted);
    if (k2) return k2;
  }

  const tryStrings = [];
  if (typeof addrObjOrString === "string") tryStrings.push(addrObjOrString);
  if (typeof fallbackString === "string") tryStrings.push(fallbackString);
  if (projectObj?.name) tryStrings.push(projectObj.name);

  for (const s of tryStrings) {
    const k = extractStreetKey_(s);
    if (k) return k;
  }
  return "";
}

function extractStreetKey_(s) {
  const raw = normalizeAddress_(String(s || ""));
  const m = raw.match(
    /\b(\d{1,6})\s+([A-Z0-9]+\s+){0,6}(ST|AVE|RD|DR|LN|CT|PL|TER|CIR|BLVD|PKWY|HWY)\b/
  );
  if (m) return m[0].trim();

  const m2 = raw.match(/\b(\d{1,6})\s+([A-Z0-9]+)(\s+[A-Z0-9]+){0,3}\b/);
  return m2 ? m2[0].trim() : "";
}

function extractStreetNumber_(s) {
  const raw = normalizeAddress_(String(s || ""));
  const m = raw.match(/\b(\d{1,6})\b/);
  return m ? m[1] : "";
}

function extractStreetStem_(s) {
  const raw = normalizeAddress_(String(s || ""));
  const k = extractStreetKey_(raw);
  const base = k || raw;

  let x = base.replace(/^\s*\d{1,6}\s+/, "").trim();
  x = x.replace(/\b(ST|AVE|RD|DR|LN|CT|PL|TER|CIR|BLVD|PKWY|HWY)\b.*$/i, "").trim();

  const parts = x.split(/\s+/).filter(Boolean);
  if (!parts.length) return "";
  return parts.slice(0, 2).join(" ");
}

function formatAddressObjOneLine_(addrObj) {
  if (!addrObj) return "";
  if (typeof addrObj === "string") return addrObj;

  const street1 = addrObj.street_address_1 || "";
  const street2 = addrObj.street_address_2 ? " " + addrObj.street_address_2 : "";
  const city = addrObj.city || "";
  const state = addrObj.state || "";
  const postal = addrObj.postal_code || "";

  const csz = [city, [state, postal].filter(Boolean).join(" ")].filter(Boolean).join(", ");
  return [(street1 + street2).trim(), csz].filter(Boolean).join(", ").trim();
}

/**************** NAME + PHONE HELPERS ****************/

function extractNameFormats_(title) {
  const parts = String(title || "").split(",");
  if (parts.length < 2) return { lastFirst: title, firstLast: title };

  const lastName = parts[0].trim();
  const firstName = parts.slice(1).join(",").trim();
  return { lastFirst: `${lastName}, ${firstName}`, firstLast: `${firstName} ${lastName}` };
}

function parsePhone_(desc) {
  const text = String(desc || "");

  const mLine = text.match(/^\s*phone\s*:\s*([^\n\r]+)/im);
  if (mLine && mLine[1]) {
    const normalized = normalizePhone_(mLine[1]);
    if (normalized) return normalized;
  }

  const m2 = text.match(/(\+?1?[\s\-\.]?\(?\d{3}\)?[\s\-\.]?\d{3}[\s\-\.]?\d{4})/);
  if (m2 && m2[1]) {
    const normalized = normalizePhone_(m2[1]);
    if (normalized) return normalized;
  }

  return "";
}

/**************** PACKET FUNCTIONS (MEASURES ONLY) ****************/

function createPacketFromTemplate_({ customerName, displayName, phone, address, datetime }) {
  const templateFile = DriveApp.getFileById(TEMPLATE_FILE_ID);
  const city = extractCityFromAddress_(address);

  const newName = city ? `${displayName} - ${city}` : `${displayName}`;
  const copy = templateFile.makeCopy(newName);
  copy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

  const ss = SpreadsheetApp.openById(copy.getId());
  const sheet = ss.getSheets()[0];

  sheet.getRange(CELL_NAME).setValue(customerName || "");
  sheet.getRange(CELL_ADDRESS).setValue(address || "");
  sheet.getRange(CELL_PHONE).setValue(phone || "");
  sheet.getRange(CELL_DATETIME).setValue(datetime);

  return { fileId: copy.getId(), url: ss.getUrl() };
}

function updateExistingPacket_(fileId, { customerName, phone, address, datetime }) {
  const ss = SpreadsheetApp.openById(fileId);
  const sheet = ss.getSheets()[0];

  sheet.getRange(CELL_NAME).setValue(customerName || "");
  sheet.getRange(CELL_ADDRESS).setValue(address || "");
  sheet.getRange(CELL_PHONE).setValue(phone || "");
  sheet.getRange(CELL_DATETIME).setValue(datetime);
}

function extractCityFromAddress_(address) {
  const parts = (address || "").split(",");
  if (parts.length < 2) return "";

  const cityStateZip = parts[1].trim();
  const m = cityStateZip.match(/^(.+?)\s+[A-Z]{2}\b/);
  if (m) return m[1].trim();

  return cityStateZip.split(/\s+/)[0].trim();
}

/**************** URL HELPERS ****************/

function buildDriveOpenUrl_(fileId) {
  return `https://drive.google.com/open?id=${fileId}`;
}

/**************** ADDRESS NORMALIZATION ****************/

function normalizeAddress_(s) {
  if (!s) return "";
  let x = String(s).toUpperCase();

  x = x.replace(/[.,]/g, " ");
  x = x.replace(/\s+/g, " ").trim();

  x = x
    .replace(/\bSTREET\b/g, "ST")
    .replace(/\bAVENUE\b/g, "AVE")
    .replace(/\bROAD\b/g, "RD")
    .replace(/\bDRIVE\b/g, "DR")
    .replace(/\bLANE\b/g, "LN")
    .replace(/\bCOURT\b/g, "CT")
    .replace(/\bPLACE\b/g, "PL")
    .replace(/\bTERRACE\b/g, "TER")
    .replace(/\bCIRCLE\b/g, "CIR")
    .replace(/\bBOULEVARD\b/g, "BLVD")
    .replace(/\bPARKWAY\b/g, "PKWY")
    .replace(/\bHIGHWAY\b/g, "HWY");

  x = x
    .replace(/\bNORTH\b/g, "N")
    .replace(/\bSOUTH\b/g, "S")
    .replace(/\bEAST\b/g, "E")
    .replace(/\bWEST\b/g, "W");

  return x;
}
