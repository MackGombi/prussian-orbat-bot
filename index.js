/***********************************************************************
 * WORLDWIDE TIMEZONE SYSTEM
 *
 * STANDARD DISPLAY RANGE: F6:F20
 * STANDARD STORAGE RANGE: G6:G20
 * GARNISON KOMPANIE DISPLAY RANGE: F6:F36
 * GARNISON KOMPANIE STORAGE RANGE: G6:G36
 *
 * Column F displays the current timezone abbreviation:
 *   PDT, PST, EDT, EST, CET, CEST, AEDT, AEST, etc.
 *
 * Column G stores the permanent IANA timezone:
 *   America/Los_Angeles
 *   Europe/Berlin
 *   Australia/Sydney
 *
 * The hidden IANA value allows the script to update existing entries
 * automatically when daylight-saving time begins or ends.
 ***********************************************************************/


/***********************************************************************
 * CONFIGURATION
 ***********************************************************************/

const TIMEZONE_CONFIG = {
  companySheets: [
    "Ostpreussisches Jäger-Bataillon Command",
    "1. Kompanie der Jäger",
    "2. Kompanie der Jäger",
    "1. Krümper-Kompanie",
    "2. Krümper-Kompanie",
    "Garnison Kompanie",
  ],

  firstRow: 6,
  lastRow: 26,

  // Column F
  displayColumn: 6,

  // Column G
  storageColumn: 7,

  refreshFunctionName: "refreshAllTimezones"
};



/***********************************************************************
 * REGIMENT SPREADSHEETS
 *
 * Approved regiment spreadsheet IDs used by the Discord bot and
 * timezone maintenance functions.
 ***********************************************************************/

const ORBAT_SPREADSHEET_IDS = [
  "1YymsOxdnXyuADCknvIwZLefnleD17_xQrSo-jE0eMnY",
  "1WJvb3A_f55TnF6ugBrpT1AjS9brlUm7SolPuCFU7g2A",
  "12947z7gDGUAZakgG5EZKkoKYYTf-gIo716ax3dnlOBk",
  "1Q5j3gYqwboiL-kHvAVnGvKI-csDuLsxZ4MjHCL9THEo"
];

function isApprovedSpreadsheetId(spreadsheetId) {
  const cleanedId = String(spreadsheetId || "").trim();

  return ORBAT_SPREADSHEET_IDS
    .map(function(id) {
      return String(id || "").trim();
    })
    .filter(function(id) {
      return id !== "" && !id.startsWith("PASTE_");
    })
    .includes(cleanedId);
}

function forEachOrbatSpreadsheet(callback) {
  ORBAT_SPREADSHEET_IDS.forEach(function(spreadsheetId) {
    const cleanedId = String(spreadsheetId || "").trim();

    if (
      cleanedId === "" ||
      cleanedId.startsWith("PASTE_")
    ) {
      return;
    }

    const spreadsheet = SpreadsheetApp.openById(cleanedId);
    callback(spreadsheet, cleanedId);
  });
}

/***********************************************************************
 * RETURNS THE ALLOWED ROW RANGE FOR EACH SHEET
 *
 * JÄGER COMPANY SHEETS:
 *   Row 6      = Company Commander
 *   Rows 7-15  = 1. Platoon
 *   Row 16     = separator
 *   Rows 17-26 = 2. Platoon
 *
 *   Timezone display/storage therefore covers F6:F26 and G6:G26.
 *
 * GARNISON KOMPANIE:
 *   No platoon structure.
 *   Uses F6:F36 and G6:G36.
 *
 * OTHER CONFIGURED SHEETS:
 *   Use the standard TIMEZONE_CONFIG member range.
 ***********************************************************************/

function getSheetRowRange(sheetName) {
  /*
   * Ostpreussisches Jäger-Bataillon Command
   * Timezone system only applies to F6:F7 / G6:G7.
   */
  if (sheetName === "Ostpreussisches Jäger-Bataillon Command") {
    return {
      firstRow: 6,
      lastRow: 7
    };
  }

  /*
   * Ostpreußisches Jäger-Bataillon company sheets.
   */
  if (
    sheetName === "1. Kompanie der Jäger" ||
    sheetName === "2. Kompanie der Jäger"
  ) {
    return {
      firstRow: 6,
      lastRow: 26
    };
  }

  /*
   * Garnison has no platoons and has additional personnel rows.
   */
  if (sheetName === "Garnison Kompanie") {
    return {
      firstRow: 6,
      lastRow: 36
    };
  }

  return {
    firstRow: TIMEZONE_CONFIG.firstRow,
    lastRow: TIMEZONE_CONFIG.lastRow
  };
}


/***********************************************************************
 * RUNS WHEN SOMEONE EDITS THE SPREADSHEET
 ***********************************************************************/

function onEdit(e) {
  if (!e || !e.range) return;

  const editedRange = e.range;
  const sheet = editedRange.getSheet();
  const sheetName = sheet.getName();

  if (!TIMEZONE_CONFIG.companySheets.includes(sheetName)) return;

  const displayColumn = TIMEZONE_CONFIG.displayColumn;
  const rowConfig = getSheetRowRange(sheetName);
  const firstAllowedRow = rowConfig.firstRow;
  const lastAllowedRow = rowConfig.lastRow;

  const editedFirstColumn = editedRange.getColumn();
  const editedLastColumn =
    editedFirstColumn + editedRange.getNumColumns() - 1;

  /*
   * Stop unless the edited range includes column F.
   *
   * This supports:
   * - Editing one cell in column F
   * - Clearing several cells in column F
   * - Pasting several timezone values into column F
   * - Pasting a larger range that includes column F
   */
  if (
    displayColumn < editedFirstColumn ||
    displayColumn > editedLastColumn
  ) {
    return;
  }

  const editedFirstRow = editedRange.getRow();
  const editedLastRow =
    editedFirstRow + editedRange.getNumRows() - 1;

  /*
   * Work only with the portion of the edit that overlaps the allowed
   * range for this sheet. Standard sheets use F6:F20, while
   * Garnison Kompanie uses F6:F36.
   */
  const firstRow = Math.max(
    editedFirstRow,
    firstAllowedRow
  );

  const lastRow = Math.min(
    editedLastRow,
    lastAllowedRow
  );

  if (firstRow > lastRow) return;

  const numberOfRows = lastRow - firstRow + 1;

  const displayRange = sheet.getRange(
    firstRow,
    displayColumn,
    numberOfRows,
    1
  );

  const storageRange = sheet.getRange(
    firstRow,
    TIMEZONE_CONFIG.storageColumn,
    numberOfRows,
    1
  );

  const enteredValues = displayRange.getDisplayValues();
  const storageValues = [];
  const displayValues = [];
  const notes = [];

  for (let index = 0; index < numberOfRows; index++) {
    const rawInput = String(
      enteredValues[index][0] || ""
    ).trim();

    /*
     * If a timezone is deleted from column F, delete the matching
     * stored IANA timezone from column G and remove any note.
     */
    if (rawInput === "") {
      storageValues.push([""]);
      displayValues.push([""]);
      notes.push([""]);
      continue;
    }

    const result = resolveTimezoneInput(rawInput);

    if (!result.success) {
      storageValues.push([""]);

      /*
       * Keep the unrecognized text visible so the user can correct it.
       */
      displayValues.push([rawInput]);

      notes.push([
        result.message ||
        "Timezone not recognized. Enter a city, country, state, province, " +
        "timezone abbreviation, UTC/GMT offset, or IANA timezone."
      ]);

      continue;
    }

    storageValues.push([result.storageValue]);
    displayValues.push([result.displayValue]);
    notes.push([""]);
  }

  /*
   * Update all affected rows in one batch.
   */
  storageRange.setValues(storageValues);
  displayRange.setValues(displayValues);
  displayRange.setNotes(notes);
}


/***********************************************************************
 * ADDS A MENU TO THE SPREADSHEET
 ***********************************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Timezone Tools")
    .addItem("Refresh all timezones", "refreshAllTimezones")
    .addItem("Install daily refresh", "installDailyTimezoneRefresh")
    .addItem("Install edit triggers for all regiments", "installAllTimezoneEditTriggers")
    .addItem("Hide timezone storage column", "hideTimezoneStorageColumns")
    .addToUi();
}


/***********************************************************************
 * RESOLVES THE USER'S INPUT
 ***********************************************************************/

function resolveTimezoneInput(rawInput) {
  const normalized = normalizeTimezoneInput(rawInput);

  /*
   * 1. US and Canadian timezone families.
   *
   * PST, PDT, and PT all resolve to America/Los_Angeles.
   * The displayed result is corrected based on the current date.
   */
  const daylightAwareAbbreviations = {
    // Pacific
    "pst": "America/Los_Angeles",
    "pdt": "America/Los_Angeles",
    "pt": "America/Los_Angeles",
    "pacific": "America/Los_Angeles",
    "pacifictime": "America/Los_Angeles",

    // Mountain
    "mst": "America/Denver",
    "mdt": "America/Denver",
    "mt": "America/Denver",
    "mountain": "America/Denver",
    "mountaintime": "America/Denver",

    // Central
    "cst": "America/Chicago",
    "cdt": "America/Chicago",
    "ct": "America/Chicago",
    "central": "America/Chicago",
    "centraltime": "America/Chicago",

    // Eastern
    "est": "America/New_York",
    "edt": "America/New_York",
    "et": "America/New_York",
    "eastern": "America/New_York",
    "easterntime": "America/New_York",

    // Alaska
    "akst": "America/Anchorage",
    "akdt": "America/Anchorage",
    "akt": "America/Anchorage",
    "alaskatime": "America/Anchorage",

    // Atlantic
    "ast": "America/Halifax",
    "adt": "America/Halifax",
    "at": "America/Halifax",
    "atlantictime": "America/Halifax",

    // Newfoundland
    "nst": "America/St_Johns",
    "ndt": "America/St_Johns",
    "nt": "America/St_Johns",
    "newfoundlandtime": "America/St_Johns",

    // Hawaii
    "hst": "Pacific/Honolulu",
    "hawaiitime": "Pacific/Honolulu"
  };

  if (daylightAwareAbbreviations[normalized]) {
    return createIanaResult(
      daylightAwareAbbreviations[normalized]
    );
  }

  /*
   * 2. Common worldwide abbreviations.
   *
   * Some abbreviations are ambiguous internationally. These shortcuts
   * use the most common regional interpretation for this spreadsheet.
   * A city or IANA timezone is more precise.
   */
  const worldwideAbbreviations = {
    "gmt": "Europe/London",
    "bst": "Europe/London",

    "wet": "Europe/Lisbon",
    "west": "Europe/Lisbon",

    "cet": "Europe/Berlin",
    "cest": "Europe/Berlin",

    "eet": "Europe/Helsinki",
    "eest": "Europe/Helsinki",

    "msk": "Europe/Moscow",

    "gst": "Asia/Dubai",

    "pkt": "Asia/Karachi",

    "ist": "Asia/Kolkata",

    "npt": "Asia/Kathmandu",

    "bdt": "Asia/Dhaka",

    "ict": "Asia/Bangkok",

    "sgt": "Asia/Singapore",

    "hkt": "Asia/Hong_Kong",

    "jst": "Asia/Tokyo",

    "kst": "Asia/Seoul",

    "wib": "Asia/Jakarta",
    "wita": "Asia/Makassar",
    "wit": "Asia/Jayapura",

    "awst": "Australia/Perth",

    "acst": "Australia/Adelaide",
    "acdt": "Australia/Adelaide",

    "aest": "Australia/Sydney",
    "aedt": "Australia/Sydney",

    "nzst": "Pacific/Auckland",
    "nzdt": "Pacific/Auckland",

    "chast": "Pacific/Chatham",
    "chadt": "Pacific/Chatham",

    "sast": "Africa/Johannesburg",

    "eat": "Africa/Nairobi",

    "wat": "Africa/Lagos",

    "cat": "Africa/Harare",

    "art": "America/Argentina/Buenos_Aires",

    "brt": "America/Sao_Paulo",

    "clt": "America/Santiago",
    "clst": "America/Santiago",

    "cot": "America/Bogota",

    "pet": "America/Lima",

    "vet": "America/Caracas",

    "bot": "America/La_Paz",

    "pyt": "America/Asuncion",
    "pyst": "America/Asuncion",

    "uyt": "America/Montevideo"
  };

  if (worldwideAbbreviations[normalized]) {
    return createIanaResult(
      worldwideAbbreviations[normalized]
    );
  }

  /*
   * 3. UTC and GMT.
   */
  if (normalized === "utc") {
    return {
      success: true,
      storageValue: "OFFSET:+00:00",
      displayValue: "UTC"
    };
  }

  /*
   * 4. UTC/GMT offsets.
   *
   * Accepted examples:
   * UTC+5
   * UTC+05:30
   * GMT-8
   * GMT + 9
   * +02:00
   */
  const offsetResult = parseUtcOffset(rawInput);

  if (offsetResult.success) {
    return offsetResult;
  }

  /*
   * 5. Any valid IANA timezone worldwide.
   *
   * Examples:
   * America/Sao_Paulo
   * Africa/Johannesburg
   * Asia/Manila
   * Pacific/Auckland
   */
  if (looksLikeIanaTimezone(rawInput)) {
    const cleanedIana = cleanIanaTimezone(rawInput);

    if (isValidIanaTimezone(cleanedIana)) {
      return createIanaResult(cleanedIana);
    }

    return {
      success: false,
      message:
        '"' + rawInput + '" does not appear to be a valid IANA timezone.'
    };
  }

  /*
   * 6. Country, region, state, province, and city aliases.
   */
  const locationAliases = getWorldwideLocationAliases();
  const ianaTimezone = locationAliases[normalized];

  if (ianaTimezone) {
    return createIanaResult(ianaTimezone);
  }

  /*
   * 7. Combined city, state/province, and country input.
   *
   * Accepted examples:
   * Berlin, Germany
   * Berlin Germany
   * Sacramento, California
   * Sacramento California USA
   * Sydney / Australia
   * Tokyo - Japan
   *
   * The function checks each meaningful part of the entry. When all
   * recognized parts point to the same timezone, that timezone is used.
   */
  const compoundLocationResult = resolveCompoundLocationInput(
    rawInput,
    locationAliases
  );

  if (compoundLocationResult.success) {
    return compoundLocationResult;
  }

  if (compoundLocationResult.recognizedButConflicting) {
    return compoundLocationResult;
  }

  /*
   * 8. Countries that span multiple timezones.
   *
   * These are deliberately not assigned one default timezone.
   */
  const ambiguousCountries = {
    "unitedstates":
      "The United States has multiple timezones. Enter a state, city, " +
      "US timezone abbreviation, or IANA timezone.",

    "usa":
      "The United States has multiple timezones. Enter a state, city, " +
      "US timezone abbreviation, or IANA timezone.",

    "us":
      "The United States has multiple timezones. Enter a state, city, " +
      "US timezone abbreviation, or IANA timezone.",

    "canada":
      "Canada has multiple timezones. Enter a province, city, or IANA timezone.",

    "mexico":
      "Mexico has multiple timezones. Enter a state, city, or IANA timezone.",

    "brazil":
      "Brazil has multiple timezones. Enter a city, state, or IANA timezone.",

    "russia":
      "Russia has multiple timezones. Enter a city, region, or IANA timezone.",

    "australia":
      "Australia has multiple timezones. Enter a state, city, or IANA timezone.",

    "indonesia":
      "Indonesia has multiple timezones. Enter a city, island, or IANA timezone.",

    "kazakhstan":
      "Enter a city or IANA timezone for Kazakhstan.",

    "chile":
      "Chile includes multiple timezone regions. Enter a city or IANA timezone.",

    "ecuador":
      "Ecuador and the Galápagos Islands use different timezones. " +
      "Enter a city or IANA timezone.",

    "newzealand":
      "New Zealand and the Chatham Islands use different timezones. " +
      "Enter a city or IANA timezone.",

    "spain":
      "Mainland Spain and the Canary Islands use different timezones. " +
      "Enter a city or region.",

    "portugal":
      "Mainland Portugal and the Azores use different timezones. " +
      "Enter a city or region.",

    "micronesia":
      "Micronesia spans multiple timezones. Enter an island or IANA timezone.",

    "democraticrepublicofthecongo":
      "The Democratic Republic of the Congo spans two timezones. " +
      "Enter a city or IANA timezone.",

    "drc":
      "The Democratic Republic of the Congo spans two timezones. " +
      "Enter a city or IANA timezone."
  };

  if (ambiguousCountries[normalized]) {
    return {
      success: false,
      message: ambiguousCountries[normalized]
    };
  }

  return {
    success: false,
    message:
      'Could not recognize "' + rawInput + '". Try a major city, state, ' +
      "province, UTC/GMT offset, or IANA timezone such as " +
      "America/Sao_Paulo, Europe/Berlin, Africa/Nairobi, " +
      "Asia/Manila, or Australia/Sydney."
  };
}


/***********************************************************************
 * RESOLVES COMBINED LOCATION INPUT
 *
 * Supports city + state/province + country combinations using commas,
 * slashes, pipes, hyphens, or ordinary spaces.
 ***********************************************************************/

function resolveCompoundLocationInput(rawInput, locationAliases) {
  const originalText = String(rawInput || "").trim();

  if (!originalText) {
    return {
      success: false
    };
  }

  /*
   * Convert common separators into spaces, then build every contiguous
   * word combination. This allows entries such as:
   *
   * "New York, United States"
   * "Sacramento California USA"
   * "Berlin / Germany"
   */
  const words = originalText
    .replace(/[|,/;]+/g, " ")
    .replace(/\s+-\s+/g, " ")
    .trim()
    .split(/\s+/)
    .filter(function(word) {
      return word !== "";
    });

  if (words.length < 2) {
    return {
      success: false
    };
  }

  const matches = [];
  const seenKeys = {};

  /*
   * Check longer phrases first so "new york city" is preferred over
   * shorter fragments such as "new york".
   */
  for (let phraseLength = words.length; phraseLength >= 1; phraseLength--) {
    for (
      let startIndex = 0;
      startIndex + phraseLength <= words.length;
      startIndex++
    ) {
      const phrase = words
        .slice(startIndex, startIndex + phraseLength)
        .join(" ");

      const normalizedPhrase = normalizeTimezoneInput(phrase);

      if (
        !normalizedPhrase ||
        seenKeys[normalizedPhrase] ||
        !locationAliases[normalizedPhrase]
      ) {
        continue;
      }

      seenKeys[normalizedPhrase] = true;

      matches.push({
        key: normalizedPhrase,
        timezone: locationAliases[normalizedPhrase],
        wordCount: phraseLength
      });
    }
  }

  /*
   * Also check whether known aliases form the beginning or end of the
   * fully normalized entry. This helps with entries where punctuation
   * was omitted, such as "BerlinGermany".
   */
  const normalizedFullInput = normalizeTimezoneInput(originalText);

  Object.keys(locationAliases).forEach(function(aliasKey) {
    if (
      aliasKey.length < 3 ||
      seenKeys[aliasKey] ||
      aliasKey === normalizedFullInput
    ) {
      return;
    }

    if (
      normalizedFullInput.startsWith(aliasKey) ||
      normalizedFullInput.endsWith(aliasKey)
    ) {
      seenKeys[aliasKey] = true;

      matches.push({
        key: aliasKey,
        timezone: locationAliases[aliasKey],
        wordCount: 1
      });
    }
  });

  if (matches.length === 0) {
    return {
      success: false
    };
  }

  const uniqueTimezones = [];

  matches.forEach(function(match) {
    if (!uniqueTimezones.includes(match.timezone)) {
      uniqueTimezones.push(match.timezone);
    }
  });

  if (uniqueTimezones.length === 1) {
    return createIanaResult(uniqueTimezones[0]);
  }

  /*
   * More than one recognized location pointed to different timezones.
   * Do not guess; ask the user for a more precise entry.
   */
  return {
    success: false,
    recognizedButConflicting: true,
    message:
      '"' + rawInput + '" contains locations that point to different ' +
      "timezones. Enter a more specific city and country, or use an " +
      "IANA timezone such as America/Los_Angeles."
  };
}


/***********************************************************************
 * WORLDWIDE LOCATION ALIASES
 *
 * This includes countries, capitals, major cities, US states,
 * Canadian provinces, Australian states, and common regions.
 *
 * Any valid IANA timezone also works even when a city is not listed.
 ***********************************************************************/

function getWorldwideLocationAliases() {
  return {
    /*******************************************************************
     * NORTH AMERICA — UNITED STATES
     *******************************************************************/

    "california": "America/Los_Angeles",
    "losangeles": "America/Los_Angeles",
    "sacramento": "America/Los_Angeles",
    "sanfrancisco": "America/Los_Angeles",
    "sandiego": "America/Los_Angeles",

    "washington": "America/Los_Angeles",
    "washingtonstate": "America/Los_Angeles",
    "seattle": "America/Los_Angeles",

    "oregon": "America/Los_Angeles",
    "portland": "America/Los_Angeles",

    "nevada": "America/Los_Angeles",
    "lasvegas": "America/Los_Angeles",

    "colorado": "America/Denver",
    "denver": "America/Denver",

    "utah": "America/Denver",
    "saltlakecity": "America/Denver",

    "newmexico": "America/Denver",
    "albuquerque": "America/Denver",

    "montana": "America/Denver",
    "wyoming": "America/Denver",

    "arizona": "America/Phoenix",
    "phoenix": "America/Phoenix",

    "texas": "America/Chicago",
    "dallas": "America/Chicago",
    "houston": "America/Chicago",
    "austin": "America/Chicago",
    "sanantonio": "America/Chicago",

    "illinois": "America/Chicago",
    "chicago": "America/Chicago",

    "wisconsin": "America/Chicago",
    "minnesota": "America/Chicago",
    "minneapolis": "America/Chicago",
    "iowa": "America/Chicago",
    "missouri": "America/Chicago",
    "kansas": "America/Chicago",
    "oklahoma": "America/Chicago",
    "arkansas": "America/Chicago",
    "louisiana": "America/Chicago",
    "neworleans": "America/Chicago",
    "mississippi": "America/Chicago",
    "alabama": "America/Chicago",
    "tennessee": "America/Chicago",
    "nashville": "America/Chicago",
    "nebraska": "America/Chicago",
    "northdakota": "America/Chicago",
    "southdakota": "America/Chicago",

    "newyork": "America/New_York",
    "newyorkcity": "America/New_York",

    "florida": "America/New_York",
    "miami": "America/New_York",
    "orlando": "America/New_York",

    "georgia": "America/New_York",
    "atlanta": "America/New_York",

    "pennsylvania": "America/New_York",
    "philadelphia": "America/New_York",
    "pittsburgh": "America/New_York",

    "newjersey": "America/New_York",
    "virginia": "America/New_York",
    "westvirginia": "America/New_York",
    "northcarolina": "America/New_York",
    "southcarolina": "America/New_York",
    "ohio": "America/New_York",
    "michigan": "America/Detroit",
    "detroit": "America/Detroit",
    "massachusetts": "America/New_York",
    "boston": "America/New_York",
    "connecticut": "America/New_York",
    "rhodeisland": "America/New_York",
    "vermont": "America/New_York",
    "newhampshire": "America/New_York",
    "maine": "America/New_York",
    "maryland": "America/New_York",
    "delaware": "America/New_York",
    "washingtondc": "America/New_York",

    "indiana": "America/Indiana/Indianapolis",
    "indianapolis": "America/Indiana/Indianapolis",

    "kentucky": "America/Kentucky/Louisville",
    "louisville": "America/Kentucky/Louisville",

    "idaho": "America/Boise",
    "boise": "America/Boise",

    "alaska": "America/Anchorage",
    "anchorage": "America/Anchorage",

    "hawaii": "Pacific/Honolulu",
    "honolulu": "Pacific/Honolulu",

    /*******************************************************************
     * CANADA
     *******************************************************************/

    "britishcolumbia": "America/Vancouver",
    "vancouver": "America/Vancouver",

    "alberta": "America/Edmonton",
    "edmonton": "America/Edmonton",
    "calgary": "America/Edmonton",

    "saskatchewan": "America/Regina",
    "regina": "America/Regina",

    "manitoba": "America/Winnipeg",
    "winnipeg": "America/Winnipeg",

    "ontario": "America/Toronto",
    "toronto": "America/Toronto",
    "ottawa": "America/Toronto",

    "quebec": "America/Toronto",
    "montreal": "America/Toronto",

    "newbrunswick": "America/Moncton",
    "moncton": "America/Moncton",

    "novascotia": "America/Halifax",
    "halifax": "America/Halifax",

    "princeedwardisland": "America/Halifax",

    "newfoundland": "America/St_Johns",
    "newfoundlandandlabrador": "America/St_Johns",
    "stjohns": "America/St_Johns",

    "yukon": "America/Whitehorse",
    "whitehorse": "America/Whitehorse",

    "northwestterritories": "America/Yellowknife",
    "yellowknife": "America/Yellowknife",

    "nunavut": "America/Iqaluit",
    "iqaluit": "America/Iqaluit",

    /*******************************************************************
     * MEXICO, CENTRAL AMERICA, AND CARIBBEAN
     *******************************************************************/

    "mexicocity": "America/Mexico_City",
    "guadalajara": "America/Mexico_City",
    "monterrey": "America/Monterrey",
    "tijuana": "America/Tijuana",
    "cancun": "America/Cancun",

    "belize": "America/Belize",
    "belizecity": "America/Belize",

    "guatemala": "America/Guatemala",
    "guatemalacity": "America/Guatemala",

    "elsalvador": "America/El_Salvador",
    "sansalvador": "America/El_Salvador",

    "honduras": "America/Tegucigalpa",
    "tegucigalpa": "America/Tegucigalpa",

    "nicaragua": "America/Managua",
    "managua": "America/Managua",

    "costarica": "America/Costa_Rica",
    "sanjosecostarica": "America/Costa_Rica",

    "panama": "America/Panama",
    "panamacity": "America/Panama",

    "bahamas": "America/Nassau",
    "nassau": "America/Nassau",

    "cuba": "America/Havana",
    "havana": "America/Havana",

    "jamaica": "America/Jamaica",
    "kingston": "America/Jamaica",

    "haiti": "America/Port-au-Prince",
    "portauprince": "America/Port-au-Prince",

    "dominicanrepublic": "America/Santo_Domingo",
    "santodomingo": "America/Santo_Domingo",

    "puertorico": "America/Puerto_Rico",
    "sanjuan": "America/Puerto_Rico",

    "trinidadandtobago": "America/Port_of_Spain",
    "portofspain": "America/Port_of_Spain",

    "barbados": "America/Barbados",
    "bridgetown": "America/Barbados",

    "grenada": "America/Grenada",
    "dominica": "America/Dominica",
    "saintlucia": "America/St_Lucia",
    "saintvincentandthegrenadines": "America/St_Vincent",
    "antiguaandbarbuda": "America/Antigua",
    "saintkittsandnevis": "America/St_Kitts",

    /*******************************************************************
     * SOUTH AMERICA
     *******************************************************************/

    "argentina": "America/Argentina/Buenos_Aires",
    "buenosaires": "America/Argentina/Buenos_Aires",

    "saopaulo": "America/Sao_Paulo",
    "riodejaneiro": "America/Sao_Paulo",
    "brasilia": "America/Sao_Paulo",

    "manaus": "America/Manaus",
    "recife": "America/Recife",

    "colombia": "America/Bogota",
    "bogota": "America/Bogota",

    "peru": "America/Lima",
    "lima": "America/Lima",

    "venezuela": "America/Caracas",
    "caracas": "America/Caracas",

    "bolivia": "America/La_Paz",
    "lapaz": "America/La_Paz",

    "paraguay": "America/Asuncion",
    "asuncion": "America/Asuncion",

    "uruguay": "America/Montevideo",
    "montevideo": "America/Montevideo",

    "guyana": "America/Guyana",
    "georgetown": "America/Guyana",

    "suriname": "America/Paramaribo",
    "paramaribo": "America/Paramaribo",

    "frenchguiana": "America/Cayenne",
    "cayenne": "America/Cayenne",

    "santiago": "America/Santiago",
    "easterisland": "Pacific/Easter",

    "quito": "America/Guayaquil",
    "guayaquil": "America/Guayaquil",
    "galapagosislands": "Pacific/Galapagos",

    /*******************************************************************
     * WESTERN AND CENTRAL EUROPE
     *******************************************************************/

    "unitedkingdom": "Europe/London",
    "uk": "Europe/London",
    "england": "Europe/London",
    "scotland": "Europe/London",
    "wales": "Europe/London",
    "northernireland": "Europe/London",
    "london": "Europe/London",

    "ireland": "Europe/Dublin",
    "dublin": "Europe/Dublin",

    "france": "Europe/Paris",
    "paris": "Europe/Paris",
    "lyon": "Europe/Paris",
    "marseille": "Europe/Paris",

    "germany": "Europe/Berlin",
    "berlin": "Europe/Berlin",

    "netherlands": "Europe/Amsterdam",
    "amsterdam": "Europe/Amsterdam",

    "belgium": "Europe/Brussels",
    "brussels": "Europe/Brussels",

    "luxembourg": "Europe/Luxembourg",

    "switzerland": "Europe/Zurich",
    "zurich": "Europe/Zurich",

    "austria": "Europe/Vienna",
    "vienna": "Europe/Vienna",

    "italy": "Europe/Rome",
    "rome": "Europe/Rome",
    "milan": "Europe/Rome",
    "naples": "Europe/Rome",

    "mainlandspain": "Europe/Madrid",
    "madrid": "Europe/Madrid",
    "barcelona": "Europe/Madrid",
    "canaryislands": "Atlantic/Canary",

    "mainlandportugal": "Europe/Lisbon",
    "lisbon": "Europe/Lisbon",
    "madeira": "Atlantic/Madeira",
    "azores": "Atlantic/Azores",

    "denmark": "Europe/Copenhagen",
    "copenhagen": "Europe/Copenhagen",

    "norway": "Europe/Oslo",
    "oslo": "Europe/Oslo",

    "sweden": "Europe/Stockholm",
    "stockholm": "Europe/Stockholm",

    "iceland": "Atlantic/Reykjavik",
    "reykjavik": "Atlantic/Reykjavik",

    "finland": "Europe/Helsinki",
    "helsinki": "Europe/Helsinki",

    /*******************************************************************
     * CENTRAL AND EASTERN EUROPE
     *******************************************************************/

    "poland": "Europe/Warsaw",
    "warsaw": "Europe/Warsaw",

    "czechia": "Europe/Prague",
    "czechrepublic": "Europe/Prague",
    "prague": "Europe/Prague",

    "slovakia": "Europe/Bratislava",
    "bratislava": "Europe/Bratislava",

    "hungary": "Europe/Budapest",
    "budapest": "Europe/Budapest",

    "slovenia": "Europe/Ljubljana",
    "ljubljana": "Europe/Ljubljana",

    "croatia": "Europe/Zagreb",
    "zagreb": "Europe/Zagreb",

    "bosniaandherzegovina": "Europe/Sarajevo",
    "sarajevo": "Europe/Sarajevo",

    "serbia": "Europe/Belgrade",
    "belgrade": "Europe/Belgrade",

    "montenegro": "Europe/Podgorica",
    "podgorica": "Europe/Podgorica",

    "northmacedonia": "Europe/Skopje",
    "skopje": "Europe/Skopje",

    "albania": "Europe/Tirane",
    "tirana": "Europe/Tirane",

    "greece": "Europe/Athens",
    "athens": "Europe/Athens",

    "bulgaria": "Europe/Sofia",
    "sofia": "Europe/Sofia",

    "romania": "Europe/Bucharest",
    "bucharest": "Europe/Bucharest",

    "moldova": "Europe/Chisinau",
    "chisinau": "Europe/Chisinau",

    "ukraine": "Europe/Kyiv",
    "kyiv": "Europe/Kyiv",

    "belarus": "Europe/Minsk",
    "minsk": "Europe/Minsk",

    "lithuania": "Europe/Vilnius",
    "vilnius": "Europe/Vilnius",

    "latvia": "Europe/Riga",
    "riga": "Europe/Riga",

    "estonia": "Europe/Tallinn",
    "tallinn": "Europe/Tallinn",

    /*******************************************************************
     * RUSSIA AND CAUCASUS
     *******************************************************************/

    "moscow": "Europe/Moscow",
    "saintpetersburg": "Europe/Moscow",
    "kaliningrad": "Europe/Kaliningrad",
    "yekaterinburg": "Asia/Yekaterinburg",
    "omsk": "Asia/Omsk",
    "novosibirsk": "Asia/Novosibirsk",
    "krasnoyarsk": "Asia/Krasnoyarsk",
    "irkutsk": "Asia/Irkutsk",
    "yakutsk": "Asia/Yakutsk",
    "vladivostok": "Asia/Vladivostok",
    "magadan": "Asia/Magadan",
    "kamchatka": "Asia/Kamchatka",

    "georgia-country": "Asia/Tbilisi",
    "tbilisi": "Asia/Tbilisi",

    "armenia": "Asia/Yerevan",
    "yerevan": "Asia/Yerevan",

    "azerbaijan": "Asia/Baku",
    "baku": "Asia/Baku",

    /*******************************************************************
     * MIDDLE EAST
     *******************************************************************/

    "turkey": "Europe/Istanbul",
    "istanbul": "Europe/Istanbul",
    "ankara": "Europe/Istanbul",

    "cyprus": "Asia/Nicosia",
    "nicosia": "Asia/Nicosia",

    "israel": "Asia/Jerusalem",
    "jerusalem": "Asia/Jerusalem",
    "telaviv": "Asia/Jerusalem",

    "palestine": "Asia/Hebron",
    "westbank": "Asia/Hebron",
    "gaza": "Asia/Gaza",

    "lebanon": "Asia/Beirut",
    "beirut": "Asia/Beirut",

    "syria": "Asia/Damascus",
    "damascus": "Asia/Damascus",

    "jordan": "Asia/Amman",
    "amman": "Asia/Amman",

    "iraq": "Asia/Baghdad",
    "baghdad": "Asia/Baghdad",

    "iran": "Asia/Tehran",
    "tehran": "Asia/Tehran",

    "saudiarabia": "Asia/Riyadh",
    "riyadh": "Asia/Riyadh",

    "kuwait": "Asia/Kuwait",

    "bahrain": "Asia/Bahrain",

    "qatar": "Asia/Qatar",
    "doha": "Asia/Qatar",

    "unitedarabemirates": "Asia/Dubai",
    "uae": "Asia/Dubai",
    "dubai": "Asia/Dubai",
    "abudhabi": "Asia/Dubai",

    "oman": "Asia/Muscat",
    "muscat": "Asia/Muscat",

    "yemen": "Asia/Aden",
    "aden": "Asia/Aden",

    /*******************************************************************
     * AFRICA
     *******************************************************************/

    "morocco": "Africa/Casablanca",
    "casablanca": "Africa/Casablanca",
    "rabat": "Africa/Casablanca",

    "algeria": "Africa/Algiers",
    "algiers": "Africa/Algiers",

    "tunisia": "Africa/Tunis",
    "tunis": "Africa/Tunis",

    "libya": "Africa/Tripoli",
    "tripoli": "Africa/Tripoli",

    "egypt": "Africa/Cairo",
    "cairo": "Africa/Cairo",

    "sudan": "Africa/Khartoum",
    "khartoum": "Africa/Khartoum",

    "southsudan": "Africa/Juba",
    "juba": "Africa/Juba",

    "ethiopia": "Africa/Addis_Ababa",
    "addisababa": "Africa/Addis_Ababa",

    "eritrea": "Africa/Asmara",
    "asmara": "Africa/Asmara",

    "djibouti": "Africa/Djibouti",

    "somalia": "Africa/Mogadishu",
    "mogadishu": "Africa/Mogadishu",

    "kenya": "Africa/Nairobi",
    "nairobi": "Africa/Nairobi",

    "uganda": "Africa/Kampala",
    "kampala": "Africa/Kampala",

    "tanzania": "Africa/Dar_es_Salaam",
    "daressalaam": "Africa/Dar_es_Salaam",

    "rwanda": "Africa/Kigali",
    "kigali": "Africa/Kigali",

    "burundi": "Africa/Bujumbura",
    "bujumbura": "Africa/Bujumbura",

    "kinshasa": "Africa/Kinshasa",
    "lubumbashi": "Africa/Lubumbashi",

    "republicofthecongo": "Africa/Brazzaville",
    "brazzaville": "Africa/Brazzaville",

    "gabon": "Africa/Libreville",
    "libreville": "Africa/Libreville",

    "equatorialguinea": "Africa/Malabo",
    "malabo": "Africa/Malabo",

    "cameroon": "Africa/Douala",
    "yaounde": "Africa/Douala",

    "centralafricanrepublic": "Africa/Bangui",
    "bangui": "Africa/Bangui",

    "chad": "Africa/Ndjamena",
    "ndjamena": "Africa/Ndjamena",

    "nigeria": "Africa/Lagos",
    "lagos": "Africa/Lagos",
    "abuja": "Africa/Lagos",

    "niger": "Africa/Niamey",
    "niamey": "Africa/Niamey",

    "benin": "Africa/Porto-Novo",
    "portonovo": "Africa/Porto-Novo",

    "togo": "Africa/Lome",
    "lome": "Africa/Lome",

    "ghana": "Africa/Accra",
    "accra": "Africa/Accra",

    "ivorycoast": "Africa/Abidjan",
    "cotedivoire": "Africa/Abidjan",
    "abidjan": "Africa/Abidjan",

    "burkinafaso": "Africa/Ouagadougou",
    "ouagadougou": "Africa/Ouagadougou",

    "mali": "Africa/Bamako",
    "bamako": "Africa/Bamako",

    "mauritania": "Africa/Nouakchott",
    "nouakchott": "Africa/Nouakchott",

    "senegal": "Africa/Dakar",
    "dakar": "Africa/Dakar",

    "gambia": "Africa/Banjul",
    "banjul": "Africa/Banjul",

    "guineabissau": "Africa/Bissau",
    "guinea": "Africa/Conakry",
    "conakry": "Africa/Conakry",

    "sierraleone": "Africa/Freetown",
    "freetown": "Africa/Freetown",

    "liberia": "Africa/Monrovia",
    "monrovia": "Africa/Monrovia",

    "caboverde": "Atlantic/Cape_Verde",
    "capeverde": "Atlantic/Cape_Verde",

    "angola": "Africa/Luanda",
    "luanda": "Africa/Luanda",

    "zambia": "Africa/Lusaka",
    "lusaka": "Africa/Lusaka",

    "zimbabwe": "Africa/Harare",
    "harare": "Africa/Harare",

    "malawi": "Africa/Blantyre",
    "lilongwe": "Africa/Blantyre",

    "mozambique": "Africa/Maputo",
    "maputo": "Africa/Maputo",

    "namibia": "Africa/Windhoek",
    "windhoek": "Africa/Windhoek",

    "botswana": "Africa/Gaborone",
    "gaborone": "Africa/Gaborone",

    "southafrica": "Africa/Johannesburg",
    "johannesburg": "Africa/Johannesburg",
    "capetown": "Africa/Johannesburg",
    "pretoria": "Africa/Johannesburg",

    "lesotho": "Africa/Maseru",
    "maseru": "Africa/Maseru",

    "eswatini": "Africa/Mbabane",
    "swaziland": "Africa/Mbabane",
    "mbabane": "Africa/Mbabane",

    "madagascar": "Indian/Antananarivo",
    "antananarivo": "Indian/Antananarivo",

    "mauritius": "Indian/Mauritius",

    "seychelles": "Indian/Mahe",

    "comoros": "Indian/Comoro",

    /*******************************************************************
     * CENTRAL ASIA
     *******************************************************************/

    "afghanistan": "Asia/Kabul",
    "kabul": "Asia/Kabul",

    "pakistan": "Asia/Karachi",
    "islamabad": "Asia/Karachi",
    "karachi": "Asia/Karachi",

    "uzbekistan": "Asia/Tashkent",
    "tashkent": "Asia/Tashkent",

    "turkmenistan": "Asia/Ashgabat",
    "ashgabat": "Asia/Ashgabat",

    "tajikistan": "Asia/Dushanbe",
    "dushanbe": "Asia/Dushanbe",

    "kyrgyzstan": "Asia/Bishkek",
    "bishkek": "Asia/Bishkek",

    "astana": "Asia/Almaty",
    "almaty": "Asia/Almaty",

    /*******************************************************************
     * SOUTH ASIA
     *******************************************************************/

    "india": "Asia/Kolkata",
    "newdelhi": "Asia/Kolkata",
    "delhi": "Asia/Kolkata",
    "mumbai": "Asia/Kolkata",
    "kolkata": "Asia/Kolkata",
    "bangalore": "Asia/Kolkata",
    "bengaluru": "Asia/Kolkata",

    "nepal": "Asia/Kathmandu",
    "kathmandu": "Asia/Kathmandu",

    "bhutan": "Asia/Thimphu",
    "thimphu": "Asia/Thimphu",

    "bangladesh": "Asia/Dhaka",
    "dhaka": "Asia/Dhaka",

    "srilanka": "Asia/Colombo",
    "colombo": "Asia/Colombo",

    "maldives": "Indian/Maldives",
    "male": "Indian/Maldives",

    /*******************************************************************
     * EAST ASIA
     *******************************************************************/

    "china": "Asia/Shanghai",
    "beijing": "Asia/Shanghai",
    "shanghai": "Asia/Shanghai",
    "guangzhou": "Asia/Shanghai",

    "hongkong": "Asia/Hong_Kong",

    "macau": "Asia/Macau",

    "taiwan": "Asia/Taipei",
    "taipei": "Asia/Taipei",

    "mongolia": "Asia/Ulaanbaatar",
    "ulaanbaatar": "Asia/Ulaanbaatar",

    "northkorea": "Asia/Pyongyang",
    "pyongyang": "Asia/Pyongyang",

    "southkorea": "Asia/Seoul",
    "korea": "Asia/Seoul",
    "seoul": "Asia/Seoul",

    "japan": "Asia/Tokyo",
    "tokyo": "Asia/Tokyo",
    "kyoto": "Asia/Tokyo",
    "osaka": "Asia/Tokyo",

    /*******************************************************************
     * SOUTHEAST ASIA
     *******************************************************************/

    "myanmar": "Asia/Yangon",
    "burma": "Asia/Yangon",
    "yangon": "Asia/Yangon",

    "thailand": "Asia/Bangkok",
    "bangkok": "Asia/Bangkok",

    "laos": "Asia/Vientiane",
    "vientiane": "Asia/Vientiane",

    "cambodia": "Asia/Phnom_Penh",
    "phnompenh": "Asia/Phnom_Penh",

    "vietnam": "Asia/Ho_Chi_Minh",
    "hanoi": "Asia/Ho_Chi_Minh",
    "hochiminhcity": "Asia/Ho_Chi_Minh",

    "malaysia": "Asia/Kuala_Lumpur",
    "kualalumpur": "Asia/Kuala_Lumpur",

    "singapore": "Asia/Singapore",

    "brunei": "Asia/Brunei",

    "philippines": "Asia/Manila",
    "manila": "Asia/Manila",

    "jakarta": "Asia/Jakarta",
    "java": "Asia/Jakarta",
    "sumatra": "Asia/Jakarta",

    "bali": "Asia/Makassar",
    "makassar": "Asia/Makassar",

    "papuaindonesia": "Asia/Jayapura",
    "jayapura": "Asia/Jayapura",

    "timorleste": "Asia/Dili",
    "easttimor": "Asia/Dili",
    "dili": "Asia/Dili",

    /*******************************************************************
     * AUSTRALIA
     *******************************************************************/

    "westernaustralia": "Australia/Perth",
    "perth": "Australia/Perth",

    "southaustralia": "Australia/Adelaide",
    "adelaide": "Australia/Adelaide",

    "northernterritory": "Australia/Darwin",
    "darwin": "Australia/Darwin",

    "queensland": "Australia/Brisbane",
    "brisbane": "Australia/Brisbane",

    "newsouthwales": "Australia/Sydney",
    "sydney": "Australia/Sydney",

    "victoria": "Australia/Melbourne",
    "melbourne": "Australia/Melbourne",

    "tasmania": "Australia/Hobart",
    "hobart": "Australia/Hobart",

    "australiancapitalterritory": "Australia/Sydney",
    "canberra": "Australia/Sydney",

    /*******************************************************************
     * NEW ZEALAND AND PACIFIC
     *******************************************************************/

    "auckland": "Pacific/Auckland",
    "wellington": "Pacific/Auckland",
    "chathamislands": "Pacific/Chatham",

    "papuanewguinea": "Pacific/Port_Moresby",
    "portmoresby": "Pacific/Port_Moresby",

    "solomonislands": "Pacific/Guadalcanal",
    "honiara": "Pacific/Guadalcanal",

    "vanuatu": "Pacific/Efate",
    "portvila": "Pacific/Efate",

    "newcaledonia": "Pacific/Noumea",
    "noumea": "Pacific/Noumea",

    "fiji": "Pacific/Fiji",
    "suva": "Pacific/Fiji",

    "tonga": "Pacific/Tongatapu",

    "samoa": "Pacific/Apia",
    "apia": "Pacific/Apia",

    "americansamoa": "Pacific/Pago_Pago",

    "cookislands": "Pacific/Rarotonga",

    "frenchpolynesia": "Pacific/Tahiti",
    "tahiti": "Pacific/Tahiti",

    "kiribati": "Pacific/Tarawa",
    "tarawa": "Pacific/Tarawa",
    "kiritimati": "Pacific/Kiritimati",

    "tuvalu": "Pacific/Funafuti",

    "nauru": "Pacific/Nauru",

    "palau": "Pacific/Palau",

    "marshallislands": "Pacific/Majuro",
    "majuro": "Pacific/Majuro",

    "guam": "Pacific/Guam",

    "northernmarianaislands": "Pacific/Saipan",

    "pohnpei": "Pacific/Pohnpei",
    "kosrae": "Pacific/Kosrae",
    "chuuk": "Pacific/Chuuk"
  };
}


/***********************************************************************
 * CREATES A SUCCESSFUL IANA RESULT
 ***********************************************************************/

function createIanaResult(ianaTimezone) {
  if (!isValidIanaTimezone(ianaTimezone)) {
    return {
      success: false,
      message:
        '"' + ianaTimezone + '" is not a valid supported timezone.'
    };
  }

  return {
    success: true,
    storageValue: "IANA:" + ianaTimezone,
    displayValue: getCurrentTimezoneLabel(ianaTimezone)
  };
}


/***********************************************************************
 * RETURNS THE CURRENT ABBREVIATION
 ***********************************************************************/

function getCurrentTimezoneLabel(ianaTimezone) {
  const now = new Date();
  const offset = Utilities.formatDate(now, ianaTimezone, "Z");

  /*
   * Apps Script can sometimes return GMT offsets instead of familiar
   * abbreviations. These custom rules produce consistent labels.
   */
  const customRules = {
    // North America
    "America/Los_Angeles": {
      "-0800": "PST",
      "-0700": "PDT"
    },

    "America/Vancouver": {
      "-0800": "PST",
      "-0700": "PDT"
    },

    "America/Denver": {
      "-0700": "MST",
      "-0600": "MDT"
    },

    "America/Boise": {
      "-0700": "MST",
      "-0600": "MDT"
    },

    "America/Phoenix": {
      "-0700": "MST"
    },

    "America/Chicago": {
      "-0600": "CST",
      "-0500": "CDT"
    },

    "America/Winnipeg": {
      "-0600": "CST",
      "-0500": "CDT"
    },

    "America/Regina": {
      "-0600": "CST"
    },

    "America/New_York": {
      "-0500": "EST",
      "-0400": "EDT"
    },

    "America/Detroit": {
      "-0500": "EST",
      "-0400": "EDT"
    },

    "America/Toronto": {
      "-0500": "EST",
      "-0400": "EDT"
    },

    "America/Indiana/Indianapolis": {
      "-0500": "EST",
      "-0400": "EDT"
    },

    "America/Kentucky/Louisville": {
      "-0500": "EST",
      "-0400": "EDT"
    },

    "America/Halifax": {
      "-0400": "AST",
      "-0300": "ADT"
    },

    "America/St_Johns": {
      "-0330": "NST",
      "-0230": "NDT"
    },

    "America/Anchorage": {
      "-0900": "AKST",
      "-0800": "AKDT"
    },

    "Pacific/Honolulu": {
      "-1000": "HST"
    },

    // Europe
    "Europe/London": {
      "+0000": "GMT",
      "+0100": "BST"
    },

    "Europe/Dublin": {
      "+0000": "GMT",
      "+0100": "IST"
    },

    "Europe/Lisbon": {
      "+0000": "WET",
      "+0100": "WEST"
    },

    "Atlantic/Canary": {
      "+0000": "WET",
      "+0100": "WEST"
    },

    "Atlantic/Azores": {
      "-0100": "AZOT",
      "+0000": "AZOST"
    },

    "Europe/Berlin": {
      "+0100": "CET",
      "+0200": "CEST"
    },

    "Europe/Paris": {
      "+0100": "CET",
      "+0200": "CEST"
    },

    "Europe/Madrid": {
      "+0100": "CET",
      "+0200": "CEST"
    },

    "Europe/Rome": {
      "+0100": "CET",
      "+0200": "CEST"
    },

    "Europe/Warsaw": {
      "+0100": "CET",
      "+0200": "CEST"
    },

    "Europe/Prague": {
      "+0100": "CET",
      "+0200": "CEST"
    },

    "Europe/Athens": {
      "+0200": "EET",
      "+0300": "EEST"
    },

    "Europe/Helsinki": {
      "+0200": "EET",
      "+0300": "EEST"
    },

    "Europe/Kyiv": {
      "+0200": "EET",
      "+0300": "EEST"
    },

    "Europe/Moscow": {
      "+0300": "MSK"
    },

    // Asia
    "Asia/Dubai": {
      "+0400": "GST"
    },

    "Asia/Karachi": {
      "+0500": "PKT"
    },

    "Asia/Kolkata": {
      "+0530": "IST"
    },

    "Asia/Kathmandu": {
      "+0545": "NPT"
    },

    "Asia/Dhaka": {
      "+0600": "BDT"
    },

    "Asia/Bangkok": {
      "+0700": "ICT"
    },

    "Asia/Ho_Chi_Minh": {
      "+0700": "ICT"
    },

    "Asia/Jakarta": {
      "+0700": "WIB"
    },

    "Asia/Singapore": {
      "+0800": "SGT"
    },

    "Asia/Shanghai": {
      "+0800": "CST (CHINA)"
    },

    "Asia/Hong_Kong": {
      "+0800": "HKT"
    },

    "Asia/Taipei": {
      "+0800": "CST (TAIWAN)"
    },

    "Asia/Makassar": {
      "+0800": "WITA"
    },

    "Asia/Tokyo": {
      "+0900": "JST"
    },

    "Asia/Seoul": {
      "+0900": "KST"
    },

    "Asia/Jayapura": {
      "+0900": "WIT"
    },

    // Australia
    "Australia/Perth": {
      "+0800": "AWST"
    },

    "Australia/Darwin": {
      "+0930": "ACST"
    },

    "Australia/Adelaide": {
      "+0930": "ACST",
      "+1030": "ACDT"
    },

    "Australia/Brisbane": {
      "+1000": "AEST"
    },

    "Australia/Sydney": {
      "+1000": "AEST",
      "+1100": "AEDT"
    },

    "Australia/Melbourne": {
      "+1000": "AEST",
      "+1100": "AEDT"
    },

    "Australia/Hobart": {
      "+1000": "AEST",
      "+1100": "AEDT"
    },

    // New Zealand
    "Pacific/Auckland": {
      "+1200": "NZST",
      "+1300": "NZDT"
    },

    "Pacific/Chatham": {
      "+1245": "CHAST",
      "+1345": "CHADT"
    },

    // Africa
    "Africa/Johannesburg": {
      "+0200": "SAST"
    },

    "Africa/Nairobi": {
      "+0300": "EAT"
    },

    "Africa/Lagos": {
      "+0100": "WAT"
    },

    // South America
    "America/Argentina/Buenos_Aires": {
      "-0300": "ART"
    },

    "America/Sao_Paulo": {
      "-0300": "BRT"
    },

    "America/Bogota": {
      "-0500": "COT"
    },

    "America/Lima": {
      "-0500": "PET"
    },

    "America/Caracas": {
      "-0400": "VET"
    },

    "America/La_Paz": {
      "-0400": "BOT"
    },

    "America/Montevideo": {
      "-0300": "UYT"
    }
  };

  const rules = customRules[ianaTimezone];

  if (rules && rules[offset]) {
    return rules[offset];
  }

  /*
   * First fallback: timezone abbreviation supplied by Apps Script.
   */
  const automaticLabel = Utilities.formatDate(
    now,
    ianaTimezone,
    "z"
  );

  if (
    automaticLabel &&
    automaticLabel !== ianaTimezone
  ) {
    return automaticLabel.toUpperCase();
  }

  /*
   * Final fallback: readable UTC offset.
   */
  return formatOffsetAsUtc(offset);
}


/***********************************************************************
 * VALIDATES AN IANA TIMEZONE
 ***********************************************************************/

function isValidIanaTimezone(ianaTimezone) {
  try {
    Utilities.formatDate(
      new Date(),
      ianaTimezone,
      "yyyy-MM-dd HH:mm:ss"
    );

    return true;
  } catch (error) {
    return false;
  }
}


/***********************************************************************
 * CHECKS WHETHER INPUT LOOKS LIKE AN IANA TIMEZONE
 ***********************************************************************/

function looksLikeIanaTimezone(input) {
  return String(input).includes("/");
}


/***********************************************************************
 * CLEANS IANA INPUT
 *
 * Allows:
 * america/los_angeles
 * America/Los_Angeles
 ***********************************************************************/

function cleanIanaTimezone(input) {
  const trimmed = String(input).trim();

  const knownPrefixes = {
    "africa": "Africa",
    "america": "America",
    "antarctica": "Antarctica",
    "arctic": "Arctic",
    "asia": "Asia",
    "atlantic": "Atlantic",
    "australia": "Australia",
    "europe": "Europe",
    "indian": "Indian",
    "pacific": "Pacific",
    "etc": "Etc"
  };

  const parts = trimmed.split("/");

  if (parts.length < 2) return trimmed;

  const prefix = knownPrefixes[parts[0].toLowerCase()];

  if (prefix) {
    parts[0] = prefix;
  }

  return parts.join("/");
}


/***********************************************************************
 * PARSES UTC/GMT OFFSETS
 ***********************************************************************/

function parseUtcOffset(input) {
  let normalized = String(input)
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "");

  // Allow a bare +05:30 or -08:00.
  if (/^[+-]/.test(normalized)) {
    normalized = "UTC" + normalized;
  }

  const match = normalized.match(
    /^(UTC|GMT)([+-])(\d{1,2})(?::?(\d{2}))?$/
  );

  if (!match) {
    return {
      success: false
    };
  }

  const sign = match[2];
  const hours = Number(match[3]);
  const minutes = match[4] ? Number(match[4]) : 0;

  if (hours > 14 || minutes > 59) {
    return {
      success: false,
      message: "The UTC/GMT offset is outside the valid range."
    };
  }

  if (hours === 14 && minutes !== 0) {
    return {
      success: false,
      message: "UTC offsets greater than UTC+14:00 are not supported."
    };
  }

  const formattedOffset =
    sign +
    String(hours).padStart(2, "0") +
    ":" +
    String(minutes).padStart(2, "0");

  const display =
    formattedOffset === "+00:00" ||
    formattedOffset === "-00:00"
      ? "UTC"
      : "UTC" + formattedOffset;

  return {
    success: true,
    storageValue: "OFFSET:" + formattedOffset,
    displayValue: display
  };
}


/***********************************************************************
 * FORMATS AN APPS SCRIPT OFFSET
 *
 * Converts +0530 to UTC+05:30.
 ***********************************************************************/

function formatOffsetAsUtc(offset) {
  if (!offset || !/^[+-]\d{4}$/.test(offset)) {
    return "UTC";
  }

  const sign = offset.substring(0, 1);
  const hours = offset.substring(1, 3);
  const minutes = offset.substring(3, 5);

  if (hours === "00" && minutes === "00") {
    return "UTC";
  }

  return "UTC" + sign + hours + ":" + minutes;
}


/***********************************************************************
 * NORMALIZES CITY, COUNTRY, AND REGION INPUT
 ***********************************************************************/

function normalizeTimezoneInput(input) {
  return String(input)
    .trim()
    .toLowerCase()

    // Remove accents:
    // São Paulo becomes sao paulo.
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")

    // Treat & as "and".
    .replace(/&/g, "and")

    // Remove punctuation and spaces.
    .replace(/['’.,()_-]/g, "")
    .replace(/\s+/g, "");
}


/***********************************************************************
 * REFRESHES ALL TIMEZONES
 *
 * This is the function run by the daily trigger.
 ***********************************************************************/

function refreshAllTimezones() {
  forEachOrbatSpreadsheet(function(spreadsheet) {
    TIMEZONE_CONFIG.companySheets.forEach(function(sheetName) {
      const sheet = spreadsheet.getSheetByName(sheetName);

      if (!sheet) return;

      const rowConfig = getSheetRowRange(sheetName);
      const numberOfRows =
        rowConfig.lastRow - rowConfig.firstRow + 1;

      const displayRange = sheet.getRange(
        rowConfig.firstRow,
        TIMEZONE_CONFIG.displayColumn,
        numberOfRows,
        1
      );

      const storageRange = sheet.getRange(
        rowConfig.firstRow,
        TIMEZONE_CONFIG.storageColumn,
        numberOfRows,
        1
      );

      const storageValues = storageRange.getValues();
      const displayValues = displayRange.getValues();
      const updatedDisplayValues = [];

      for (let index = 0; index < numberOfRows; index++) {
        const storedValue = String(
          storageValues[index][0] || ""
        ).trim();

        let displayValue = displayValues[index][0];

        if (storedValue.startsWith("IANA:")) {
          const ianaTimezone = storedValue.substring(5);

          if (isValidIanaTimezone(ianaTimezone)) {
            displayValue =
              getCurrentTimezoneLabel(ianaTimezone);
          }
        } else if (storedValue.startsWith("OFFSET:")) {
          const offset = storedValue.substring(7);

          displayValue =
            offset === "+00:00" ||
            offset === "-00:00"
              ? "UTC"
              : "UTC" + offset;
        }

        updatedDisplayValues.push([displayValue]);
      }

      displayRange.setValues(updatedDisplayValues);
    });
  });
}


/***********************************************************************
 * INSTALLS EDIT TRIGGERS FOR ALL REGIMENT SPREADSHEETS
 *
 * Run this manually once after adding all three spreadsheet IDs.
 * This preserves manual timezone entry in column F on every regiment
 * spreadsheet, not only the original spreadsheet.
 ***********************************************************************/

function installAllTimezoneEditTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "onEdit") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ORBAT_SPREADSHEET_IDS.forEach(function(spreadsheetId) {
    const cleanedId = String(spreadsheetId || "").trim();

    if (
      cleanedId === "" ||
      cleanedId.startsWith("PASTE_")
    ) {
      return;
    }

    ScriptApp.newTrigger("onEdit")
      .forSpreadsheet(cleanedId)
      .onEdit()
      .create();
  });

  SpreadsheetApp.getUi().alert(
    "Timezone edit triggers were installed for all configured regiment spreadsheets."
  );
}


/***********************************************************************
 * INSTALLS THE DAILY REFRESH TRIGGER
 *
 * Run this function manually one time.
 ***********************************************************************/

function installDailyTimezoneRefresh() {
  const functionName =
    TIMEZONE_CONFIG.refreshFunctionName;

  /*
   * Remove existing copies of the same trigger so multiple daily
   * triggers are not accidentally created.
   */
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  /*
   * Runs once daily between approximately 3:00 AM and 4:00 AM
   * according to the Apps Script project's configured timezone.
   */
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();

  refreshAllTimezones();

  SpreadsheetApp.getUi().alert(
    "Daily timezone refresh installed successfully."
  );
}


/***********************************************************************
 * HIDES COLUMN G ON ALL COMPANY SHEETS
 ***********************************************************************/

function hideTimezoneStorageColumns() {
  forEachOrbatSpreadsheet(function(spreadsheet) {
    TIMEZONE_CONFIG.companySheets.forEach(function(sheetName) {
      const sheet = spreadsheet.getSheetByName(sheetName);

      if (!sheet) return;

      sheet.hideColumns(
        TIMEZONE_CONFIG.storageColumn,
        1
      );
    });
  });

  SpreadsheetApp.getUi().alert(
    "Timezone storage column G has been hidden in all configured regiment spreadsheets."
  );
}


/***********************************************************************
 * PROCESSES A SINGLE ROW FOR THE DISCORD BOT
 ***********************************************************************/
function processTimezoneRow(
  spreadsheetId,
  sheetName,
  rowNumber,
  rawInput
) {
  if (!isApprovedSpreadsheetId(spreadsheetId)) {
    return {
      success: false,
      message:
        "The selected regiment spreadsheet is not approved."
    };
  }

  const spreadsheet =
    SpreadsheetApp.openById(spreadsheetId);

  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    return {
      success: false,
      message: 'Worksheet not found: "' + sheetName + '".'
    };
  }

  const rowConfig = getSheetRowRange(sheetName);

  if (
    !Number.isInteger(rowNumber) ||
    rowNumber < rowConfig.firstRow ||
    rowNumber > rowConfig.lastRow
  ) {
    return {
      success: false,
      message:
        "Invalid row number. Allowed rows for this worksheet are " +
        rowConfig.firstRow +
        " through " +
        rowConfig.lastRow +
        "."
    };
  }

  const timezoneInput = String(rawInput || "").trim();

  const displayCell = sheet.getRange(
    rowNumber,
    TIMEZONE_CONFIG.displayColumn
  );

  const storageCell = sheet.getRange(
    rowNumber,
    TIMEZONE_CONFIG.storageColumn
  );

  if (timezoneInput === "") {
    displayCell.clearContent();
    displayCell.clearNote();
    storageCell.clearContent();

    return {
      success: true,
      spreadsheetId: spreadsheetId,
      sheetName: sheetName,
      row: rowNumber,
      input: "",
      displayValue: "",
      storageValue: ""
    };
  }

  const result = resolveTimezoneInput(timezoneInput);

  if (!result.success) {
    storageCell.clearContent();
    displayCell.setValue(timezoneInput);
    displayCell.setNote(
      result.message || "Timezone could not be recognized."
    );

    return {
      success: false,
      message:
        result.message || "Timezone could not be recognized."
    };
  }

  storageCell.setValue(result.storageValue);
  displayCell.setValue(result.displayValue);
  displayCell.clearNote();

  return {
    success: true,
    spreadsheetId: spreadsheetId,
    sheetName: sheetName,
    row: rowNumber,
    input: timezoneInput,
    displayValue: result.displayValue,
    storageValue: result.storageValue
  };
}



/***********************************************************************
 * RETURNS AVAILABLE COMPANY WORKSHEETS FOR DISCORD AUTOCOMPLETE
 *
 * Only worksheet tabs that:
 *   1. Actually exist in the selected regiment spreadsheet, and
 *   2. Are listed in TIMEZONE_CONFIG.companySheets
 *
 * are returned. This prevents overview/dashboard tabs from appearing
 * in Discord while allowing each regiment spreadsheet to have a
 * different set of company tabs.
 ***********************************************************************/

function getCompaniesForSpreadsheet(spreadsheetId) {
  if (!isApprovedSpreadsheetId(spreadsheetId)) {
    return {
      success: false,
      message:
        "The selected regiment spreadsheet is not approved.",
      companies: []
    };
  }

  const spreadsheet =
    SpreadsheetApp.openById(spreadsheetId);

  const allowedCompanyNames =
    TIMEZONE_CONFIG.companySheets.map(function(name) {
      return String(name || "").trim();
    });

  const companies = spreadsheet
    .getSheets()
    .map(function(sheet) {
      return sheet.getName();
    })
    .filter(function(sheetName) {
      return allowedCompanyNames.includes(sheetName);
    });

  return {
    success: true,
    spreadsheetId: spreadsheetId,
    companies: companies
  };
}


/***********************************************************************
 * WEB APP CONFIGURATION
 *
 * The Discord bot sends a POST request to the deployed /exec URL.
 * The secret below must exactly match BOT_WEBHOOK_SECRET in the bot's
 * .env file.
 ***********************************************************************/

const WEBHOOK_CONFIG = {
  secret: "prussian-orbat-timezone-2026-secret"
};


/***********************************************************************
 * CREATES A JSON WEB RESPONSE
 ***********************************************************************/

function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


/***********************************************************************
 * OPTIONAL HEALTH CHECK
 *
 * Opening the deployed /exec URL in a browser should return JSON that
 * confirms that the web app is online.
 ***********************************************************************/


function doGet() {
  return createJsonResponse({
    success: true,
    message: "Prussian ORBAT timezone web app is online."
  });
}


/***********************************************************************
 * RECEIVES TIMEZONE REQUESTS FROM THE DISCORD BOT
 ***********************************************************************/

const BOT_WEBHOOK_SECRET =
  "prussian-orbat-timezone-2026-secret";
  
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({
        success: false,
        message: "No JSON request body was received."
      });
    }

    const requestBody = JSON.parse(e.postData.contents);

    console.log("Webhook request:");
    console.log(JSON.stringify(requestBody));

    if (requestBody.secret !== BOT_WEBHOOK_SECRET) {
      return createJsonResponse({
        success: false,
        message: "Invalid webhook secret."
      });
    }

    const spreadsheetId = String(
      requestBody.spreadsheetId || ""
    ).trim();

    if (!spreadsheetId) {
      return createJsonResponse({
        success: false,
        message: "A spreadsheet ID was not provided.",
        receivedRequest: requestBody
      });
    }

    if (!isApprovedSpreadsheetId(spreadsheetId)) {
      return createJsonResponse({
        success: false,
        message:
          "The selected regiment spreadsheet is not approved.",
        receivedSpreadsheetId: spreadsheetId
      });
    }

    /*
     * Dynamic company autocomplete request.
     *
     * The Discord bot sends:
     * {
     *   action: "getCompanies",
     *   secret: "...",
     *   spreadsheetId: "..."
     * }
     */
    const action = String(
      requestBody.action || ""
    ).trim();

    if (action === "getCompanies") {
      return createJsonResponse(
        getCompaniesForSpreadsheet(spreadsheetId)
      );
    }

    /*
     * Existing timezone-processing request.
     */
    const sheetName = String(
      requestBody.sheetName ||
      requestBody.company ||
      requestBody.worksheet ||
      ""
    ).trim();

    const rawRowNumber =
      requestBody.rowNumber ??
      requestBody.row ??
      requestBody.spreadsheetRow ??
      requestBody.sheetRow;

    const rowNumber = parseInt(rawRowNumber, 10);

    const timezone = String(
      requestBody.timezone ||
      requestBody.timezoneSubmitted ||
      ""
    ).trim();

    if (!sheetName) {
      return createJsonResponse({
        success: false,
        message: "A worksheet name was not provided.",
        receivedRequest: requestBody
      });
    }

    if (!Number.isInteger(rowNumber) || rowNumber < 1) {
      return createJsonResponse({
        success: false,
        message:
          "A valid spreadsheet row number was not provided.",
        receivedRowNumber: rawRowNumber,
        receivedRequest: requestBody
      });
    }

    if (!timezone) {
      return createJsonResponse({
        success: false,
        message: "A timezone was not provided.",
        receivedRequest: requestBody
      });
    }

    const result = processTimezoneRow(
      spreadsheetId,
      sheetName,
      rowNumber,
      timezone
    );

    return createJsonResponse(result);

  } catch (error) {
    console.error(error);

    return createJsonResponse({
      success: false,
      message:
        error && error.message
          ? error.message
          : String(error)
    });
  }
}