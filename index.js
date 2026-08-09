import "dotenv/config";

import fs from "node:fs";
import {
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
  Client,
  EmbedBuilder,
  Events,
  GatewayIntentBits,
  MessageFlags,
  StringSelectMenuBuilder,
  UserSelectMenuBuilder
} from "discord.js";
import { google } from "googleapis";

/*
|--------------------------------------------------------------------------
| Environment helpers
|--------------------------------------------------------------------------
*/

function cleanEnvironmentValue(value) {
  return String(value || "")
    .replace(/[\u2028\u2029\uFEFF]/g, "")
    .trim();
}

/*
|--------------------------------------------------------------------------
| Configuration
|--------------------------------------------------------------------------
*/

const DISCORD_TOKEN = cleanEnvironmentValue(
  process.env.DISCORD_TOKEN
);

const APPS_SCRIPT_URL = cleanEnvironmentValue(
  process.env.APPS_SCRIPT_URL
);

const BOT_WEBHOOK_SECRET = cleanEnvironmentValue(
  process.env.BOT_WEBHOOK_SECRET
);

const ROVER_API_KEY = cleanEnvironmentValue(
  process.env.ROVER_API_KEY
);

const ROVER_API_BASE_URL =
  cleanEnvironmentValue(
    process.env.ROVER_API_BASE_URL
  ) || "https://registry.rover.link/api";

const ERSTESSCHLESISCHES_SPREADSHEET_ID = cleanEnvironmentValue(
    process.env.ERSTESSCHLESISCHES_SPREADSHEET_ID
  );

const WESTPREUSSISCHES_SPREADSHEET_ID = cleanEnvironmentValue(
  process.env.WESTPREUSSISCHES_SPREADSHEET_ID
);

const JAEGER_SPREADSHEET_ID = cleanEnvironmentValue(
  process.env.JAEGER_SPREADSHEET_ID
);

const SCHUETZEN_SPREADSHEET_ID =
  cleanEnvironmentValue(
    process.env.SCHUETZEN_SPREADSHEET_ID
  ) || "1Q5j3gYqwboiL-kHvAVnGvKI-csDuLsxZ4MjHCL9THEo";

const SERVICE_ACCOUNT_FILE = "./service-account.json";

const GOOGLE_SERVICE_ACCOUNT_BASE64 =
  cleanEnvironmentValue(
    process.env.GOOGLE_SERVICE_ACCOUNT_BASE64
  );

const FIRST_MEMBER_ROW = 6;
const COMMAND_LAST_MEMBER_ROW = 7;
const STANDARD_LAST_MEMBER_ROW = 20;
const GARNISON_LAST_MEMBER_ROW = 36;
const GARNISON_SHEET_NAME = "Garnison Kompanie";

const SCHUETZEN_POSITIONS = {
  "company_commander": {
    displayName: "Company Commander",
    type: "command",
    firstRow: 6,
    lastRow: 6
  },
  "1_platoon": {
    displayName: "1. Platoon",
    type: "platoon",
    firstRow: 7,
    lastRow: 16
  },
  "2_platoon": {
    displayName: "2. Platoon",
    type: "platoon",
    firstRow: 17,
    lastRow: 26
  }
};

const SCHUETZEN_PLATOONS = {
  "1_platoon":
    SCHUETZEN_POSITIONS["1_platoon"],
  "2_platoon":
    SCHUETZEN_POSITIONS["2_platoon"]
};

const TWO_ROW_COMMAND_SHEETS = new Set([
  "Ostpreussisches Jäger-Bataillon Command",
  "6. Generalstab",
  "11. Generalstab"
]);
const ORBAT_LOG_CHANNEL_NAME = "orbat-logs";

/*
|--------------------------------------------------------------------------
| Regiment configuration
|--------------------------------------------------------------------------
|
| The command's regiment option value can use any of the aliases below.
| The first value in each aliases array is recommended for deploy-commands.js.
|--------------------------------------------------------------------------
*/

const REGIMENTS = [
  {
    key: "erstes_schlesisches",
    displayName:
      "11. Erstes Schlesisches Infanterie-Regiment",
    nicknamePrefix: "11",
    spreadsheetId:
      ERSTESSCHLESISCHES_SPREADSHEET_ID,
    aliases: [
      "11_schlesisches",
      "erstes_schlesisches",
      "11. erstes schlesisches infanterie-regiment",
      "11. schlesisches",
      "schlesisches",
      "11th silesian"
    ]
  },
  {
    key: "westpreussisches",
    displayName:
      "6. Erstes Westpreußisches Infanterie-Regiment",
    nicknamePrefix: "6",
    spreadsheetId:
      WESTPREUSSISCHES_SPREADSHEET_ID,
    aliases: [
      "6_westpreussisches",
      "westpreussisches",
      "westpreußisches",
      "6. westpreussisches",
      "6. westpreußisches",
      "6. erstes westpreußisches infanterie-regiment",
      "6. erstes westpreussisches infanterie-regiment",
      "6th west prussian",
      "west prussian"
    ]
  },
  {
    key: "jaeger",
    displayName: "Ostpreußisches Jäger-Bataillon",
    nicknamePrefix: "Ost",
    spreadsheetId: JAEGER_SPREADSHEET_ID,
    aliases: [
      "ostpreussisches_jaeger",
      "jaeger",
      "jäger",
      "ostpreussisches jaeger-bataillon",
      "ostpreußisches jäger-bataillon",
      "ostpreussisches jäger-bataillon",
      "ostpreußisches jaeger-bataillon",
      "ostpreussisches jaeger",
      "ostpreußisches jäger"
    ]
  },
  {
    key: "schuetzen",
    displayName: "Schlesisches Schützen-Bataillon",
    nicknamePrefix: "Sch",
    spreadsheetId: SCHUETZEN_SPREADSHEET_ID,
    aliases: [
      "schlesisches_schuetzen",
      "schuetzen",
      "schützen",
      "schlesisches schützen-bataillon",
      "schlesisches schuetzen-bataillon",
      "schlesisches schützen",
      "schlesisches schuetzen"
    ]
  }
];

/*
|--------------------------------------------------------------------------
| Validation
|--------------------------------------------------------------------------
*/

const requiredEnvironmentValues = [
  ["DISCORD_TOKEN", DISCORD_TOKEN],
  ["APPS_SCRIPT_URL", APPS_SCRIPT_URL],
  ["BOT_WEBHOOK_SECRET", BOT_WEBHOOK_SECRET],
  ["ROVER_API_KEY", ROVER_API_KEY],
  [
    "ERSTESSCHLESISCHES_SPREADSHEET_ID",
    ERSTESSCHLESISCHES_SPREADSHEET_ID
  ],
  [
    "WESTPREUSSISCHES_SPREADSHEET_ID",
    WESTPREUSSISCHES_SPREADSHEET_ID
  ],
  ["JAEGER_SPREADSHEET_ID", JAEGER_SPREADSHEET_ID],
  ["SCHUETZEN_SPREADSHEET_ID", SCHUETZEN_SPREADSHEET_ID]
];

for (const [name, value] of requiredEnvironmentValues) {
  if (!value) {
    throw new Error(`${name} is missing from the .env file.`);
  }
}

if (!APPS_SCRIPT_URL.endsWith("/exec")) {
  throw new Error(
    "APPS_SCRIPT_URL must be the deployed Apps Script web-app URL ending in /exec."
  );
}

if (
  !GOOGLE_SERVICE_ACCOUNT_BASE64 &&
  !fs.existsSync(SERVICE_ACCOUNT_FILE)
) {
  throw new Error(
    "Google credentials are missing. Add service-account.json locally or set GOOGLE_SERVICE_ACCOUNT_BASE64 when hosting."
  );
}


/*
|--------------------------------------------------------------------------
| RoVer and Roblox account lookup
|--------------------------------------------------------------------------
*/

function extractRobloxId(data) {
  const possibleValues = [
    data?.robloxId,
    data?.roblox_id,
    data?.robloxUserId,
    data?.roblox_user_id,
    data?.userId,
    data?.user_id,
    data?.id,
    data?.user?.robloxId,
    data?.user?.roblox_id,
    data?.user?.id,
    data?.roblox?.id
  ];

  for (const value of possibleValues) {
    const normalized = String(value || "").trim();

    if (/^\d+$/.test(normalized)) {
      return normalized;
    }
  }

  return null;
}

function extractRobloxUsername(data) {
  const possibleValues = [
    data?.robloxUsername,
    data?.roblox_username,
    data?.username,
    data?.name,
    data?.user?.robloxUsername,
    data?.user?.roblox_username,
    data?.user?.username,
    data?.roblox?.username,
    data?.roblox?.name
  ];

  for (const value of possibleValues) {
    const normalized = String(value || "").trim();

    if (normalized) {
      return normalized;
    }
  }

  return null;
}

async function getRobloxUsernameById(robloxId) {
  const response = await fetch(
    `https://users.roblox.com/v1/users/${encodeURIComponent(
      robloxId
    )}`,
    {
      headers: {
        Accept: "application/json"
      }
    }
  );

  if (!response.ok) {
    throw new Error(
      `Roblox username lookup failed with HTTP ${response.status}.`
    );
  }

  const data = await response.json();
  const username = String(data?.name || "").trim();

  if (!username) {
    throw new Error(
      "Roblox returned an account without a username."
    );
  }

  return username;
}

async function getVerifiedRobloxAccount({
  guildId,
  discordId
}) {
  const endpoint =
    `${ROVER_API_BASE_URL.replace(/\/+$/, "")}` +
    `/guilds/${encodeURIComponent(guildId)}` +
    `/discord-to-roblox/${encodeURIComponent(discordId)}`;

  const response = await fetch(endpoint, {
    method: "GET",
    headers: {
      Accept: "application/json",
      Authorization: `Bearer ${ROVER_API_KEY}`
    }
  });

  let data = null;

  try {
    data = await response.json();
  } catch {
    data = null;
  }

  if (response.status === 404) {
    throw new Error(
      "ROVER_NOT_VERIFIED"
    );
  }

  if (
    response.status === 401 ||
    response.status === 403
  ) {
    throw new Error(
      "ROVER_ACCESS_DENIED"
    );
  }

  if (!response.ok) {
    const roverMessage = String(
      data?.message ||
      data?.error ||
      ""
    ).trim();

    throw new Error(
      roverMessage
        ? `RoVer API error: ${roverMessage}`
        : `RoVer API request failed with HTTP ${response.status}.`
    );
  }

  const robloxId = extractRobloxId(data);
  let robloxUsername =
    extractRobloxUsername(data);

  if (!robloxId && !robloxUsername) {
    throw new Error(
      "ROVER_INVALID_RESPONSE"
    );
  }

  if (!robloxUsername && robloxId) {
    robloxUsername =
      await getRobloxUsernameById(
        robloxId
      );
  }

  return {
    robloxId,
    robloxUsername
  };
}

/*
|--------------------------------------------------------------------------
| Discord client
|--------------------------------------------------------------------------
*/

const client = new Client({
  intents: [GatewayIntentBits.Guilds]
});

/*
|--------------------------------------------------------------------------
| Google Sheets authentication
|--------------------------------------------------------------------------
*/

function loadGoogleCredentials() {
  if (GOOGLE_SERVICE_ACCOUNT_BASE64) {
    try {
      const decodedJson = Buffer
        .from(
          GOOGLE_SERVICE_ACCOUNT_BASE64,
          "base64"
        )
        .toString("utf8");

      return JSON.parse(decodedJson);
    } catch (error) {
      throw new Error(
        "GOOGLE_SERVICE_ACCOUNT_BASE64 could not be decoded as valid service-account JSON."
      );
    }
  }

  return JSON.parse(
    fs.readFileSync(
      SERVICE_ACCOUNT_FILE,
      "utf8"
    )
  );
}

const auth = new google.auth.GoogleAuth({
  credentials: loadGoogleCredentials(),
  scopes: [
    "https://www.googleapis.com/auth/spreadsheets"
  ]
});

const sheets = google.sheets({
  version: "v4",
  auth
});

/*
|--------------------------------------------------------------------------
| Utility functions
|--------------------------------------------------------------------------
*/

function normalizeText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .toLowerCase()
    .replace(/[’']/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function resolveRegiment(regimentValue) {
  const normalizedValue = normalizeText(regimentValue);

  const regiment = REGIMENTS.find(entry => {
    const searchableValues = [
      entry.key,
      entry.displayName,
      ...entry.aliases
    ];

    return searchableValues.some(
      value => normalizeText(value) === normalizedValue
    );
  });

  if (!regiment) {
    throw new Error("UNKNOWN_REGIMENT");
  }

  return regiment;
}


function isSchuetzenRegiment(
  regimentOrSpreadsheetId
) {
  const spreadsheetId =
    typeof regimentOrSpreadsheetId === "string"
      ? regimentOrSpreadsheetId
      : regimentOrSpreadsheetId?.spreadsheetId;

  return (
    String(spreadsheetId || "").trim() ===
    SCHUETZEN_SPREADSHEET_ID
  );
}

function resolveSchuetzenPosition(
  positionValue
) {
  const normalized =
    normalizeText(positionValue);

  const entries =
    Object.entries(
      SCHUETZEN_POSITIONS
    );

  for (
    const [key, config] of entries
  ) {
    const aliases = [
      key,
      config.displayName,
      key.replace(/_/g, " "),
      config.displayName.replace(".", "")
    ];

    if (
      aliases.some(
        alias =>
          normalizeText(alias) ===
          normalized
      )
    ) {
      return {
        key,
        ...config
      };
    }
  }

  throw new Error(
    "INVALID_SCHUETZEN_POSITION"
  );
}

function resolveSchuetzenPlatoon(
  platoonValue
) {
  const position =
    resolveSchuetzenPosition(
      platoonValue
    );

  if (position.type !== "platoon") {
    throw new Error(
      "INVALID_SCHUETZEN_PLATOON"
    );
  }

  return position;
}

function getSchuetzenPositionForRow(
  row
) {
  for (
    const [key, config] of
    Object.entries(
      SCHUETZEN_POSITIONS
    )
  ) {
    if (
      row >= config.firstRow &&
      row <= config.lastRow
    ) {
      return {
        key,
        ...config
      };
    }
  }

  return null;
}

function getSchuetzenPlatoonForRow(
  row
) {
  const position =
    getSchuetzenPositionForRow(
      row
    );

  return position?.type === "platoon"
    ? position
    : null;
}

function getLastMemberRowForSpreadsheet(
  spreadsheetId,
  sheetName
) {
  if (
    isSchuetzenRegiment(
      spreadsheetId
    )
  ) {
    return 26;
  }

  return getLastMemberRow(
    sheetName
  );
}

function escapeSheetName(sheetName) {
  return `'${String(sheetName).replaceAll("'", "''")}'`;
}

function getLastMemberRow(sheetName) {
  const cleanedSheetName =
    String(sheetName || "").trim();

  if (TWO_ROW_COMMAND_SHEETS.has(cleanedSheetName)) {
    return COMMAND_LAST_MEMBER_ROW;
  }

  if (cleanedSheetName === GARNISON_SHEET_NAME) {
    return GARNISON_LAST_MEMBER_ROW;
  }

  return STANDARD_LAST_MEMBER_ROW;
}

/*
|--------------------------------------------------------------------------
| ORBAT rank ordering and Krümper rules
|--------------------------------------------------------------------------
|
| Lower priority number = higher position on the sheet.
|--------------------------------------------------------------------------
*/

const RANK_SORT_PRIORITY = new Map([
  ["oberst", 1],
  ["oberst lieutenant", 2],
  ["major", 3],
  ["stabs kapitan", 4],
  ["kapitan", 5],
  ["premier lieutenant", 6],
  ["sekonde lieutenant", 7],
  ["fahnrich", 8],
  ["feldwebel", 9],
  ["sergeant", 10],
  ["korporal", 11],
  ["vizekorporal", 12],
  ["obergefreiter", 13],
  ["gefreiter", 14],
  ["obersoldat", 15],
  ["soldat", 16],
  ["rekrut", 17]
]);

function isRekrutRank(rank) {
  return normalizeText(rank) === "rekrut";
}

function isFirstKrumperCompany(sheetName) {
  return (
    normalizeText(sheetName) ===
    normalizeText("1. Krümper-Kompanie")
  );
}

function getCurrentOrbatDate() {
  return new Intl.DateTimeFormat(
    "en-US",
    {
      timeZone: "America/Los_Angeles",
      month: "2-digit",
      day: "2-digit",
      year: "numeric"
    }
  ).format(new Date());
}

async function writeKrumperEntryDate({
  spreadsheetId,
  sheetName,
  row
}) {
  if (!isFirstKrumperCompany(sheetName)) {
    return;
  }

  const safeSheetName =
    escapeSheetName(sheetName);

  const dateValue =
    getCurrentOrbatDate();

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range:
      `${safeSheetName}!H${row}`,
    valueInputOption: "RAW",
    requestBody: {
      values: [[dateValue]]
    }
  });

  console.log(
    "KRUMPER ENTRY DATE WRITTEN:",
    {
      sheetName,
      row,
      date: dateValue
    }
  );
}

function isSecondKrumperCompany(sheetName) {
  return (
    normalizeText(sheetName) ===
    normalizeText("2. Krümper-Kompanie")
  );
}

function isMusketierCompany(sheetName) {
  return normalizeText(sheetName).includes("musketier");
}

/*
|--------------------------------------------------------------------------
| Attendance rules
|--------------------------------------------------------------------------
|
| Standard eligible companies:
| Monday H, Tuesday I, Wednesday J, Thursday K,
| Friday L, Saturday M, Sunday N.
|
| 2. Krümper-Kompanie:
| Saturday H, Sunday I only.
|--------------------------------------------------------------------------
*/

const ATTENDANCE_DAYS = [
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday",
  "Sunday"
];

const STANDARD_ATTENDANCE_COLUMNS = new Map([
  ["monday", "H"],
  ["tuesday", "I"],
  ["wednesday", "J"],
  ["thursday", "K"],
  ["friday", "L"],
  ["saturday", "M"],
  ["sunday", "N"]
]);

const SECOND_KRUMPER_ATTENDANCE_COLUMNS = new Map([
  ["saturday", "H"],
  ["sunday", "I"]
]);

/*
 * These are the Discord attendance-status dropdown options.
 * If your Google Sheets dropdown uses different exact wording,
 * only change these values here and in deploy-commands.js.
 */
const ATTENDANCE_STATUS_CHOICES = [
  "DM",
  "RSVP",
  "MAYB",
  "NO",
  "PRES",
  "EXC",
  "AWOL",
  "LEFT"
];

function isGeneralstabOrCommandSheet(sheetName) {
  const normalized =
    normalizeText(sheetName);

  return (
    normalized.includes("generalstab") ||
    TWO_ROW_COMMAND_SHEETS.has(
      String(sheetName || "").trim()
    )
  );
}

function isAttendanceExcludedCompany(sheetName) {
  return (
    isGeneralstabOrCommandSheet(sheetName) ||
    normalizeText(sheetName) ===
      normalizeText(GARNISON_SHEET_NAME) ||
    isFirstKrumperCompany(sheetName)
  );
}

function getAttendanceColumn(sheetName, day) {
  const normalizedDay =
    normalizeText(day);

  if (isAttendanceExcludedCompany(sheetName)) {
    throw new Error(
      "ATTENDANCE_NOT_ALLOWED_FOR_COMPANY"
    );
  }

  if (isSecondKrumperCompany(sheetName)) {
    const column =
      SECOND_KRUMPER_ATTENDANCE_COLUMNS.get(
        normalizedDay
      );

    if (!column) {
      throw new Error(
        "SECOND_KRUMPER_WEEKEND_ONLY"
      );
    }

    return column;
  }

  const column =
    STANDARD_ATTENDANCE_COLUMNS.get(
      normalizedDay
    );

  if (!column) {
    throw new Error(
      "INVALID_ATTENDANCE_DAY"
    );
  }

  return column;
}


async function getMemberAttendanceSummary({
  spreadsheetId,
  sheetName,
  row
}) {
  const safeSheetName = escapeSheetName(sheetName);

  if (isFirstKrumperCompany(sheetName)) {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${safeSheetName}!H${row}`
    });
    const entryDate = String(response.data.values?.[0]?.[0] || "").trim();
    return { lines: [`**Entry Date:** ${entryDate || "Blank"}`] };
  }

  if (isAttendanceExcludedCompany(sheetName)) {
    return { lines: ["**Attendance:** Not tracked for this company."] };
  }

  if (isSecondKrumperCompany(sheetName)) {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${safeSheetName}!H${row}:I${row}`
    });
    const values = response.data.values?.[0] || [];
    return {
      lines: [
        `**Saturday:** ${String(values[0] || "").trim() || "Blank"}`,
        `**Sunday:** ${String(values[1] || "").trim() || "Blank"}`
      ]
    };
  }

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${safeSheetName}!H${row}:N${row}`
  });
  const values = response.data.values?.[0] || [];
  return {
    lines: ATTENDANCE_DAYS.map(
      (day, i) => `**${day}:** ${String(values[i] || "").trim() || "Blank"}`
    )
  };
}

async function getCompanyRosterSummary({
  spreadsheetId,
  sheetName
}) {
  const safeSheetName = escapeSheetName(sheetName);
  const lastMemberRow = getLastMemberRow(sheetName);

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${safeSheetName}!C${FIRST_MEMBER_ROW}:E${lastMemberRow}`,
    majorDimension: "ROWS"
  });

  const members = (response.data.values || [])
    .map((row, index) => ({
      robloxUsername: String(row?.[0] || "").trim(),
      discordId: String(row?.[1] || "").trim(),
      rank: String(row?.[2] || "").trim(),
      row: FIRST_MEMBER_ROW + index
    }))
    .filter(m => m.robloxUsername || m.discordId || m.rank);

  const rankCounts = new Map();
  for (const member of members) {
    const rank = member.rank || "Unknown";
    rankCounts.set(rank, (rankCounts.get(rank) || 0) + 1);
  }

  const rankBreakdown = [...rankCounts.entries()].sort((a, b) => {
    const rankA = RANK_SORT_PRIORITY.get(normalizeText(a[0])) ?? 999;
    const rankB = RANK_SORT_PRIORITY.get(normalizeText(b[0])) ?? 999;
    return rankA !== rankB
      ? rankA - rankB
      : a[0].localeCompare(b[0], undefined, { sensitivity: "base" });
  });

  return { members, rankBreakdown };
}

async function getMissingAttendanceMembers({
  spreadsheetId,
  sheetName,
  day
}) {
  const attendanceColumn = getAttendanceColumn(sheetName, day);
  const safeSheetName = escapeSheetName(sheetName);
  const lastMemberRow = getLastMemberRow(sheetName);
  const lastColumn = isSecondKrumperCompany(sheetName) ? "I" : "N";

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${safeSheetName}!C${FIRST_MEMBER_ROW}:${lastColumn}${lastMemberRow}`,
    majorDimension: "ROWS"
  });

  const attendanceIndex =
    columnLetterToNumber(attendanceColumn) - columnLetterToNumber("C");

  return (response.data.values || [])
    .map((row, index) => ({
      robloxUsername: String(row?.[0] || "").trim(),
      discordId: String(row?.[1] || "").trim(),
      rank: String(row?.[2] || "").trim(),
      attendance: String(row?.[attendanceIndex] || "").trim(),
      row: FIRST_MEMBER_ROW + index
    }))
    .filter(m =>
      (m.robloxUsername || m.discordId || m.rank) &&
      !m.attendance
    );
}


async function getCompanyAttendanceView({
  spreadsheetId,
  sheetName,
  day
}) {
  const attendanceColumn =
    getAttendanceColumn(
      sheetName,
      day
    );

  const safeSheetName =
    escapeSheetName(sheetName);

  const lastMemberRow =
    getLastMemberRow(sheetName);

  const lastColumn =
    isSecondKrumperCompany(sheetName)
      ? "I"
      : "N";

  const response =
    await sheets.spreadsheets.values.get({
      spreadsheetId,
      range:
        `${safeSheetName}!C${FIRST_MEMBER_ROW}:` +
        `${lastColumn}${lastMemberRow}`,
      majorDimension: "ROWS"
    });

  const attendanceIndex =
    columnLetterToNumber(attendanceColumn) -
    columnLetterToNumber("C");

  const members =
    (response.data.values || [])
      .map((row, index) => ({
        robloxUsername:
          String(row?.[0] || "").trim(),
        discordId:
          String(row?.[1] || "").trim(),
        rank:
          String(row?.[2] || "").trim(),
        attendance:
          String(
            row?.[attendanceIndex] || ""
          ).trim(),
        row:
          FIRST_MEMBER_ROW + index
      }))
      .filter(
        member =>
          member.robloxUsername ||
          member.discordId ||
          member.rank
      );

  const groups =
    new Map();

  for (const status of ATTENDANCE_STATUS_CHOICES) {
    groups.set(status, []);
  }

  groups.set("BLANK", []);

  for (const member of members) {
    const normalizedStatus =
      ATTENDANCE_STATUS_CHOICES.find(
        status =>
          normalizeText(status) ===
          normalizeText(member.attendance)
      );

    const bucket =
      normalizedStatus || "BLANK";

    groups.get(bucket).push(member);
  }

  return {
    members,
    groups,
    attendanceColumn
  };
}

async function getCompanyAuditRows({
  spreadsheetId,
  sheetName
}) {
  const safeSheetName =
    escapeSheetName(sheetName);

  const lastMemberRow =
    getLastMemberRow(sheetName);

  const response =
    await sheets.spreadsheets.values.get({
      spreadsheetId,
      range:
        `${safeSheetName}!C${FIRST_MEMBER_ROW}:F${lastMemberRow}`,
      majorDimension: "ROWS"
    });

  return (response.data.values || [])
    .map((row, index) => ({
      robloxUsername:
        String(row?.[0] || "").trim(),
      discordId:
        String(row?.[1] || "").trim(),
      rank:
        String(row?.[2] || "").trim(),
      timezone:
        String(row?.[3] || "").trim(),
      row:
        FIRST_MEMBER_ROW + index
    }))
    .filter(
      member =>
        member.robloxUsername ||
        member.discordId ||
        member.rank ||
        member.timezone
    );
}

async function auditOrbatRegiments(
  regiments
) {
  const issues = [];
  const discordLocations =
    new Map();
  const usernameLocations =
    new Map();

  for (const regiment of regiments) {
    const companies =
      await getCompanySheetNames(
        regiment
      );

    for (const company of companies) {
      const rows =
        await getCompanyAuditRows({
          spreadsheetId:
            regiment.spreadsheetId,
          sheetName:
            company
        });

      for (const member of rows) {
        const location =
          `${regiment.displayName} / ${company} / Row ${member.row}`;

        if (!member.robloxUsername) {
          issues.push(
            `Missing Roblox username — ${location}`
          );
        }

        if (!member.discordId) {
          issues.push(
            `Missing Discord ID — ${location}`
          );
        }

        if (!member.rank) {
          issues.push(
            `Missing rank — ${location}`
          );
        } else if (
          !RANK_SORT_PRIORITY.has(
            normalizeText(member.rank)
          )
        ) {
          issues.push(
            `Unknown rank "${member.rank}" — ${location}`
          );
        }

        if (!member.timezone) {
          issues.push(
            `Missing timezone — ${location}`
          );
        }

        if (
          isFirstKrumperCompany(
            company
          ) &&
          member.rank &&
          !isRekrutRank(
            member.rank
          )
        ) {
          issues.push(
            `Non-Rekrut in 1. Krümper-Kompanie (${member.rank}) — ${location}`
          );
        }

        if (
          !isFirstKrumperCompany(
            company
          ) &&
          isRekrutRank(
            member.rank
          )
        ) {
          issues.push(
            `Rekrut outside 1. Krümper-Kompanie — ${location}`
          );
        }

        if (member.discordId) {
          if (
            discordLocations.has(
              member.discordId
            )
          ) {
            issues.push(
              `Duplicate Discord ID ${member.discordId} — ${discordLocations.get(member.discordId)} AND ${location}`
            );
          } else {
            discordLocations.set(
              member.discordId,
              location
            );
          }
        }

        if (member.robloxUsername) {
          const normalizedUsername =
            normalizeText(
              member.robloxUsername
            );

          if (
            usernameLocations.has(
              normalizedUsername
            )
          ) {
            issues.push(
              `Duplicate Roblox username ${member.robloxUsername} — ${usernameLocations.get(normalizedUsername)} AND ${location}`
            );
          } else {
            usernameLocations.set(
              normalizedUsername,
              location
            );
          }
        }
      }
    }
  }

  return issues;
}

async function setAttendanceMarker({
  spreadsheetId,
  sheetName,
  row,
  day,
  attendance
}) {
  const column =
    getAttendanceColumn(
      sheetName,
      day
    );

  const safeSheetName =
    escapeSheetName(sheetName);

  const range =
    `${safeSheetName}!${column}${row}`;

  const previousResponse =
    await sheets.spreadsheets.values.get({
      spreadsheetId,
      range
    });

  const previousValue = String(
    previousResponse.data.values?.[0]?.[0] || ""
  ).trim();

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[attendance]]
    }
  });

  return {
    column,
    range: `${column}${row}`,
    previousValue
  };
}

function assertCompanyRankAllowed(sheetName, rank) {
  const firstKrumper =
    isFirstKrumperCompany(sheetName);

  const rekrut =
    isRekrutRank(rank);

  if (firstKrumper && !rekrut) {
    throw new Error(
      "FIRST_KRUMPER_REKRUT_ONLY"
    );
  }

  if (rekrut && !firstKrumper) {
    throw new Error(
      "REKRUT_FIRST_KRUMPER_ONLY"
    );
  }
}

async function findMemberRowInCompanyByDiscordId({
  spreadsheetId,
  sheetName,
  discordId
}) {
  const safeSheetName = escapeSheetName(sheetName);
  const lastMemberRow =
    getLastMemberRowForSpreadsheet(
      spreadsheetId,
      sheetName
    );

  const response =
    await sheets.spreadsheets.values.get({
      spreadsheetId,
      range:
        `${safeSheetName}!D${FIRST_MEMBER_ROW}:` +
        `D${lastMemberRow}`
    });

  const values = response.data.values || [];
  const targetId = String(discordId || "").trim();

  for (let i = 0; i < values.length; i += 1) {
    if (String(values[i]?.[0] || "").trim() === targetId) {
      return FIRST_MEMBER_ROW + i;
    }
  }

  return null;
}

function getSortLastColumn(sheetName) {
  /*
   * Move the full member-owned row data when rank sorting occurs.
   *
   * Normal attendance companies:
   *   C:N = identity, rank/timezone data, Monday-Sunday attendance.
   *
   * 2. Krümper-Kompanie:
   *   C:I = identity/rank data plus Saturday H and Sunday I.
   *
   * 1. Krümper-Kompanie:
   *   C:H = identity/rank data plus entry date H.
   *
   * Generalstab/command/Garnison:
   *   C:G = normal member data only.
   */
  if (isFirstKrumperCompany(sheetName)) {
    return "H";
  }

  if (isSecondKrumperCompany(sheetName)) {
    return "I";
  }

  if (isAttendanceExcludedCompany(sheetName)) {
    return "G";
  }

  return "N";
}

function columnLetterToNumber(column) {
  let result = 0;

  for (
    const character of
    String(column || "").toUpperCase()
  ) {
    result =
      result * 26 +
      (character.charCodeAt(0) - 64);
  }

  return result;
}


async function sortRowRangeByRank({
  spreadsheetId,
  sheetName,
  firstRow,
  lastRow
}) {
  const safeSheetName =
    escapeSheetName(sheetName);

  const lastColumn =
    getSortLastColumn(sheetName);

  const columnCount =
    columnLetterToNumber(lastColumn) -
    columnLetterToNumber("C") +
    1;

  const slotCount =
    lastRow - firstRow + 1;

  const response =
    await sheets.spreadsheets.values.get({
      spreadsheetId,
      range:
        `${safeSheetName}!C${firstRow}:` +
        `${lastColumn}${lastRow}`,
      majorDimension: "ROWS"
    });

  const members =
    (response.data.values || [])
      .map(row => {
        const padded =
          Array.from(
            { length: columnCount },
            (_, index) =>
              row[index] ?? ""
          );

        return {
          rowData: padded
        };
      })
      .filter(member =>
        String(
          member.rowData[0] || ""
        ).trim() ||
        String(
          member.rowData[1] || ""
        ).trim()
      );

  members.sort((a, b) => {
    const rankA =
      RANK_SORT_PRIORITY.get(
        normalizeText(
          a.rowData[2]
        )
      ) ?? 999;

    const rankB =
      RANK_SORT_PRIORITY.get(
        normalizeText(
          b.rowData[2]
        )
      ) ?? 999;

    if (rankA !== rankB) {
      return rankA - rankB;
    }

    return String(
      a.rowData[0] || ""
    ).localeCompare(
      String(
        b.rowData[0] || ""
      ),
      undefined,
      {
        sensitivity: "base"
      }
    );
  });

  const rewrittenRows =
    Array.from(
      { length: slotCount },
      (_, index) =>
        members[index]?.rowData ||
        Array(
          columnCount
        ).fill("")
    );

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range:
      `${safeSheetName}!C${firstRow}:` +
      `${lastColumn}${lastRow}`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: rewrittenRows
    }
  });

  return members.length;
}

async function sortCompanyByRank({
  spreadsheetId,
  sheetName,
  platoon = null
}) {
  if (
    isSchuetzenRegiment(
      spreadsheetId
    )
  ) {
    if (platoon) {
      const platoonConfig =
        typeof platoon === "string"
          ? resolveSchuetzenPlatoon(
              platoon
            )
          : platoon;

      return sortRowRangeByRank({
        spreadsheetId,
        sheetName,
        firstRow:
          platoonConfig.firstRow,
        lastRow:
          platoonConfig.lastRow
      });
    }

    let total = 0;

    for (
      const platoonConfig of
      Object.values(
        SCHUETZEN_PLATOONS
      )
    ) {
      total +=
        await sortRowRangeByRank({
          spreadsheetId,
          sheetName,
          firstRow:
            platoonConfig.firstRow,
          lastRow:
            platoonConfig.lastRow
        });
    }

    return total;
  }

  const safeSheetName =
    escapeSheetName(sheetName);

  const lastMemberRow =
    getLastMemberRow(sheetName);

  const slotCount =
    lastMemberRow -
    FIRST_MEMBER_ROW +
    1;

  const lastColumn =
    getSortLastColumn(sheetName);

  const columnCount =
    columnLetterToNumber(lastColumn) -
    columnLetterToNumber("C") +
    1;

  /*
   * Read every cell belonging to the member before sorting. This keeps
   * attendance attached to the correct person if a promotion changes
   * their row position.
   */
  const response =
    await sheets.spreadsheets.values.get({
      spreadsheetId,
      range:
        `${safeSheetName}!C${FIRST_MEMBER_ROW}:` +
        `${lastColumn}${lastMemberRow}`,
      majorDimension: "ROWS"
    });

  const sourceRows =
    response.data.values || [];

  const members =
    sourceRows
      .map((row, originalIndex) => {
        const padded =
          Array.from(
            { length: columnCount },
            (_, index) =>
              row[index] ?? ""
          );

        return {
          rowData: padded,
          originalIndex
        };
      })
      .filter(member =>
        String(
          member.rowData[0] || ""
        ).trim() ||
        String(
          member.rowData[1] || ""
        ).trim()
      );

  members.sort((a, b) => {
    const rankA =
      RANK_SORT_PRIORITY.get(
        normalizeText(
          a.rowData[2]
        )
      ) ?? 999;

    const rankB =
      RANK_SORT_PRIORITY.get(
        normalizeText(
          b.rowData[2]
        )
      ) ?? 999;

    if (rankA !== rankB) {
      return rankA - rankB;
    }

    /*
     * Keep the previous alphabetical behavior for members of the
     * same rank. Their full attendance record moves with them.
     */
    return String(
      a.rowData[0] || ""
    ).localeCompare(
      String(
        b.rowData[0] || ""
      ),
      undefined,
      {
        sensitivity: "base"
      }
    );
  });

  const rewrittenRows =
    Array.from(
      { length: slotCount },
      (_, index) =>
        members[index]?.rowData ||
        Array(
          columnCount
        ).fill("")
    );

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range:
      `${safeSheetName}!C${FIRST_MEMBER_ROW}:` +
      `${lastColumn}${lastMemberRow}`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: rewrittenRows
    }
  });

  return members.length;
}

async function findAvailableMusketierCompany(regiment) {
  const companies =
    await getCompanySheetNames(regiment);

  const musketierCompanies =
    companies
      .filter(isMusketierCompany)
      .sort((a, b) =>
        a.localeCompare(
          b,
          undefined,
          {
            numeric: true,
            sensitivity: "base"
          }
        )
      );

  if (musketierCompanies.length === 0) {
    throw new Error("NO_MUSKETIER_COMPANY");
  }

  for (const companyName of musketierCompanies) {
    try {
      await findFirstEmptyRow({
        spreadsheetId: regiment.spreadsheetId,
        sheetName: companyName
      });

      return companyName;
    } catch (error) {
      if (error?.message !== "COMPANY_FULL") {
        throw error;
      }
    }
  }

  throw new Error("MUSKETIER_COMPANIES_FULL");
}


async function findFirstKrumperCompany(regiment) {
  const companies =
    await getCompanySheetNames(regiment);

  const firstKrumper =
    companies.find(
      isFirstKrumperCompany
    );

  if (!firstKrumper) {
    throw new Error(
      "FIRST_KRUMPER_COMPANY_NOT_FOUND"
    );
  }

  try {
    await findFirstEmptyRow({
      spreadsheetId:
        regiment.spreadsheetId,
      sheetName:
        firstKrumper
    });
  } catch (error) {
    if (error?.message === "COMPANY_FULL") {
      throw new Error(
        "FIRST_KRUMPER_COMPANY_FULL"
      );
    }

    throw error;
  }

  return firstKrumper;
}

/*
|--------------------------------------------------------------------------
| Discord nickname formatting
|--------------------------------------------------------------------------
|
| Nickname format:
| Regiment prefix | abbreviated rank RobloxUsername
|
| Example:
| 10 | Sgt. robloxusername
|--------------------------------------------------------------------------
*/

const RANK_NICKNAME_ABBREVIATIONS = new Map([
  ["rekrut", "Rkt."],
  ["soldat", "Sdt."],
  ["obersoldat", "OSdt."],
  ["gefreiter", "Gefr."],
  ["obergefreiter", "OGefr."],
  ["vizekorporal", "VzKpl."],
  ["korporal", "Kpl."],
  ["sergeant", "Sgt."],
  ["feldwebel", "Fw."],
  ["fahnrich", "Fhr."],
  ["sekonde lieutenant", "2-Ltn."],
  ["premier lieutenant", "1-Ltn."],
  ["kapitan", "Kpt."],
  ["stabs kapitan", "S-Kpt."],
  ["major", "Maj."],
  ["oberst lieutenant", "Obs-Lt."],
  ["oberst", "Obs."]
]);

function getRankNicknameAbbreviation(rank) {
  const normalizedRank = normalizeText(rank);

  return (
    RANK_NICKNAME_ABBREVIATIONS.get(normalizedRank) ||
    String(rank || "").trim()
  );
}

function buildDiscordNickname({
  regiment,
  rank,
  robloxUsername
}) {
  const prefix =
    String(regiment?.nicknamePrefix || "").trim();

  const abbreviatedRank =
    getRankNicknameAbbreviation(rank);

  const username =
    String(robloxUsername || "").trim();

  const nickname =
    `${prefix} | ${abbreviatedRank} ${username}`.trim();

  /*
   * Discord nicknames have a maximum length of 32 characters.
   * Roblox usernames are trimmed only when necessary.
   */
  return nickname.slice(0, 32);
}

async function updateDiscordNickname({
  interaction,
  discordUserId,
  regiment,
  rank,
  robloxUsername
}) {
  if (!interaction.guild) {
    throw new Error(
      "The Discord server could not be accessed for nickname updating."
    );
  }

  const guildMember =
    await interaction.guild.members.fetch(
      discordUserId
    );

  const nickname = buildDiscordNickname({
    regiment,
    rank,
    robloxUsername
  });

  await guildMember.setNickname(
    nickname,
    "Prussian ORBAT automatic nickname update"
  );

  return nickname;
}

async function resetDiscordNicknameToRobloxUsername({
  interaction,
  discordUserId,
  robloxUsername
}) {
  if (!interaction.guild) {
    throw new Error(
      "The Discord server could not be accessed for nickname updating."
    );
  }

  const username =
    String(robloxUsername || "").trim();

  if (!username) {
    throw new Error(
      "The member does not have a Roblox username saved in the ORBAT."
    );
  }

  const guildMember =
    await interaction.guild.members.fetch(
      discordUserId
    );

  const nickname = username.slice(0, 32);

  await guildMember.setNickname(
    nickname,
    "Prussian ORBAT member removal"
  );

  return nickname;
}

/*
|--------------------------------------------------------------------------
| Google Sheets functions
|--------------------------------------------------------------------------
*/


async function findFirstEmptyRowInRange({
  spreadsheetId,
  sheetName,
  firstRow,
  lastRow
}) {
  const safeSheetName =
    escapeSheetName(sheetName);

  const response =
    await sheets.spreadsheets.values.get({
      spreadsheetId,
      range:
        `${safeSheetName}!C${firstRow}:` +
        `C${lastRow}`
    });

  const values =
    response.data.values || [];

  for (
    let row = firstRow;
    row <= lastRow;
    row += 1
  ) {
    const arrayIndex =
      row - firstRow;

    const value =
      values[arrayIndex]?.[0];

    if (
      !value ||
      String(value).trim() === ""
    ) {
      return row;
    }
  }

  throw new Error(
    "PLATOON_FULL"
  );
}

async function findFirstEmptyRow({
  spreadsheetId,
  sheetName
}) {
  const safeSheetName = escapeSheetName(sheetName);
  const lastMemberRow = getLastMemberRow(sheetName);

  const response =
    await sheets.spreadsheets.values.get({
      spreadsheetId,
      range:
        `${safeSheetName}!C${FIRST_MEMBER_ROW}:` +
        `C${lastMemberRow}`
    });

  const values = response.data.values || [];

  for (
    let row = FIRST_MEMBER_ROW;
    row <= lastMemberRow;
    row += 1
  ) {
    const arrayIndex = row - FIRST_MEMBER_ROW;
    const value = values[arrayIndex]?.[0];

    if (!value || String(value).trim() === "") {
      return row;
    }
  }

  throw new Error("COMPANY_FULL");
}

async function findMemberByDiscordId(discordId) {
  const targetDiscordId = String(discordId).trim();

  for (const regiment of REGIMENTS) {
    const companyNames =
      await getCompanySheetNames(
        regiment.spreadsheetId
      );

    if (companyNames.length === 0) {
      continue;
    }

    const ranges = companyNames.map(companyName => {
      const safeSheetName =
        escapeSheetName(companyName);

      const lastMemberRow =
        getLastMemberRowForSpreadsheet(
          regiment.spreadsheetId,
          companyName
        );

      return (
        `${safeSheetName}!D${FIRST_MEMBER_ROW}:` +
        `D${lastMemberRow}`
      );
    });

    const response =
      await sheets.spreadsheets.values.batchGet({
        spreadsheetId: regiment.spreadsheetId,
        ranges,
        majorDimension: "ROWS"
      });

    const valueRanges =
      response.data.valueRanges || [];

    for (
      let companyIndex = 0;
      companyIndex < companyNames.length;
      companyIndex += 1
    ) {
      const companyName =
        companyNames[companyIndex];

      const values =
        valueRanges[companyIndex]?.values || [];

      for (
        let rowIndex = 0;
        rowIndex < values.length;
        rowIndex += 1
      ) {
        const storedDiscordId = String(
          values[rowIndex]?.[0] || ""
        ).trim();

        if (
          storedDiscordId &&
          storedDiscordId === targetDiscordId
        ) {
          return {
            regiment,
            companyName,
            row:
              FIRST_MEMBER_ROW + rowIndex
          };
        }
      }
    }
  }

  return null;
}

async function getMemberRecord({
  spreadsheetId,
  sheetName,
  row
}) {
  const safeSheetName = escapeSheetName(sheetName);

  const response =
    await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${safeSheetName}!C${row}:F${row}`
    });

  const values = response.data.values?.[0] || [];

  return {
    robloxUsername: String(values[0] || "").trim(),
    discordId: String(values[1] || "").trim(),
    rank: String(values[2] || "").trim(),
    timezone: String(values[3] || "").trim()
  };
}

async function removeMemberFromSheet({
  spreadsheetId,
  sheetName,
  row
}) {
  const safeSheetName =
    escapeSheetName(sheetName);

  /*
   * When a member LEAVES a company, only C:F is transferable member
   * information. Attendance must stay company-specific and therefore
   * must be erased from the old company.
   *
   * Normal attendance companies:
   *   H:N = Monday-Sunday attendance -> clear all.
   *
   * 2. Krümper-Kompanie:
   *   H:I = Saturday-Sunday attendance -> clear both.
   *
   * 1. Krümper-Kompanie:
   *   H = entry date -> clear it when the member leaves.
   *
   * Generalstab / Garnison / other excluded sheets:
   *   no attendance markers are transferred.
   *
   * Column G is internal timezone storage and is cleaned by the
   * subsequent company sort; it is never copied to the destination.
   */
  const ranges = [
    `${safeSheetName}!C${row}`,
    `${safeSheetName}!D${row}`,
    `${safeSheetName}!E${row}`,
    `${safeSheetName}!F${row}`
  ];

  if (isFirstKrumperCompany(sheetName)) {
    ranges.push(
      `${safeSheetName}!H${row}`
    );
  } else if (
    isSecondKrumperCompany(sheetName)
  ) {
    ranges.push(
      `${safeSheetName}!H${row}:I${row}`
    );
  } else if (
    !isAttendanceExcludedCompany(
      sheetName
    )
  ) {
    ranges.push(
      `${safeSheetName}!H${row}:N${row}`
    );
  }

  await sheets.spreadsheets.values.batchClear({
    spreadsheetId,
    requestBody: {
      ranges
    }
  });

  console.log(
    "ORBAT ROW CLEARED:",
    {
      sheetName,
      row,
      transferredColumns: ["C", "D", "E", "F"],
      clearedRanges: ranges
    }
  );
}

async function updateMemberRank({
  spreadsheetId,
  sheetName,
  row,
  rank
}) {
  const safeSheetName = escapeSheetName(sheetName);

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${safeSheetName}!E${row}`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[rank]]
    }
  });
}

async function addMemberToSheet({
  spreadsheetId,
  sheetName,
  robloxUsername,
  discordId,
  rank,
  timezone,
  position = null,
  platoon = null
}) {
  const schuetzen =
    isSchuetzenRegiment(
      spreadsheetId
    );

  if (!schuetzen) {
    /*
     * Final hard-lock validation for the normal infantry structure.
     */
    assertCompanyRankAllowed(
      sheetName,
      rank
    );
  }

  let row;

  if (schuetzen) {
    const requestedPosition =
      position || platoon;

    if (!requestedPosition) {
      throw new Error(
        "SCHUETZEN_POSITION_REQUIRED"
      );
    }

    const positionConfig =
      typeof requestedPosition === "string"
        ? resolveSchuetzenPosition(
            requestedPosition
          )
        : requestedPosition;

    if (
      positionConfig.key ===
      "company_commander"
    ) {
      const safeSheetName =
        escapeSheetName(sheetName);

      const commanderResponse =
        await sheets.spreadsheets.values.get({
          spreadsheetId,
          range:
            `${safeSheetName}!C6:D6`
        });

      const commanderRow =
        commanderResponse.data.values?.[0] ||
        [];

      const commanderOccupied =
        String(
          commanderRow[0] || ""
        ).trim() ||
        String(
          commanderRow[1] || ""
        ).trim();

      if (commanderOccupied) {
        throw new Error(
          "SCHUETZEN_COMMANDER_OCCUPIED"
        );
      }

      row = 6;
    } else {
      row =
        await findFirstEmptyRowInRange({
          spreadsheetId,
          sheetName,
          firstRow:
            positionConfig.firstRow,
          lastRow:
            positionConfig.lastRow
        });
    }
  } else {
    row =
      await findFirstEmptyRow({
        spreadsheetId,
        sheetName
      });
  }

  const safeSheetName = escapeSheetName(sheetName);

  const writeData = [
    {
      range: `${safeSheetName}!C${row}`,
      values: [[robloxUsername]]
    },
    {
      range: `${safeSheetName}!D${row}`,
      values: [[discordId]]
    },
    {
      range: `${safeSheetName}!E${row}`,
      values: [[rank]]
    },
    {
      range: `${safeSheetName}!F${row}`,
      values: [[timezone]]
    }
  ];

  /*
   * Write the Krümper entry date in the SAME Google Sheets request as
   * the member data. This guarantees H is populated before any timezone
   * processing or rank sorting happens.
   */
  if (
    !schuetzen &&
    isFirstKrumperCompany(sheetName)
  ) {
    const entryDate =
      getCurrentOrbatDate();

    writeData.push({
      range: `${safeSheetName}!H${row}`,
      values: [[entryDate]]
    });

    console.log(
      "KRUMPER ENTRY DATE QUEUED:",
      {
        sheetName,
        row,
        date: entryDate
      }
    );
  } else if (!schuetzen) {
    /*
     * A date must NEVER follow a member into another company.
     * Explicitly blank H on every non-1. Krümper-Kompanie write in case
     * that roster slot contains stale data from an older bot version.
     */
    writeData.push({
      range: `${safeSheetName}!H${row}`,
      values: [[""]]
    });
  }

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: writeData
    }
  });

  return row;
}

/*
|--------------------------------------------------------------------------
| Apps Script webhook
|--------------------------------------------------------------------------
*/

async function processTimezoneWithAppsScript({
  spreadsheetId,
  sheetName,
  row,
  timezone
}) {
  const controller = new AbortController();

  const timeout = setTimeout(() => {
    controller.abort();
  }, 15000);

  try {
    const response = await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        secret: BOT_WEBHOOK_SECRET,
        spreadsheetId,
        sheetName,
        rowNumber: row,
        timezone
      }),
      signal: controller.signal,
      redirect: "follow"
    });

    const responseText = await response.text();

    let result;

    try {
      result = JSON.parse(responseText);
    } catch {
      throw new Error(
        "Apps Script returned an invalid response: " +
        responseText.slice(0, 300)
      );
    }

    if (!response.ok) {
      throw new Error(
        result.message ||
        `Apps Script returned HTTP ${response.status}.`
      );
    }

    if (!result.success) {
      throw new Error(
        result.message ||
        "Apps Script could not process the timezone."
      );
    }

    return result;
  } catch (error) {
    if (error?.name === "AbortError") {
      throw new Error(
        "Apps Script did not respond within 15 seconds."
      );
    }

    throw error;
  } finally {
    clearTimeout(timeout);
  }
}


/*
|--------------------------------------------------------------------------
| Dynamic company autocomplete
|--------------------------------------------------------------------------
|
| Company names are read directly from the selected regiment spreadsheet.
| Results are cached briefly so Discord autocomplete responds quickly.
|--------------------------------------------------------------------------
*/

const COMPANY_CACHE_TTL_MS = 5 * 60 * 1000;
const companyCache = new Map();

const EXCLUDED_COMPANY_SHEETS = new Set([
  "overview",
  "dashboard",
  "statistics",
  "timezone data",
  "settings",
  "configuration",
  "config",
  "summary",
  "quotas",
  "quota",
  "rankings",
  "logs",
  "archive"
]);

function isRegimentOverviewSheet(
  sheetName,
  regiment
) {
  const normalizedSheetName =
    normalizeText(sheetName);

  const regimentNames = [
    regiment.key,
    regiment.displayName,
    ...regiment.aliases
  ].map(normalizeText);

  return regimentNames.includes(
    normalizedSheetName
  );
}

async function getCompanySheetNames(
  regimentOrSpreadsheetId
) {
  const regiment =
    typeof regimentOrSpreadsheetId === "object" &&
    regimentOrSpreadsheetId !== null
      ? regimentOrSpreadsheetId
      : REGIMENTS.find(
          configuredRegiment =>
            configuredRegiment.spreadsheetId ===
            regimentOrSpreadsheetId
        );

  const spreadsheetId =
    typeof regimentOrSpreadsheetId === "string"
      ? regimentOrSpreadsheetId
      : regimentOrSpreadsheetId?.spreadsheetId;

  if (!spreadsheetId) {
    throw new Error(
      "getCompanySheetNames received no valid spreadsheet ID."
    );
  }

  const cached = companyCache.get(spreadsheetId);

  let names;

  if (
    cached &&
    Date.now() - cached.loadedAt < COMPANY_CACHE_TTL_MS
  ) {
    names = cached.names;
  } else {
    const response = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: "sheets.properties.title"
    });

    names = (response.data.sheets || [])
      .map(sheet => String(
        sheet?.properties?.title || ""
      ).trim())
      .filter(Boolean);

    companyCache.set(spreadsheetId, {
      loadedAt: Date.now(),
      names
    });
  }

  return names
    .filter(name =>
      !EXCLUDED_COMPANY_SHEETS.has(
        name.toLowerCase()
      )
    )
    .filter(name => {
      if (!regiment) {
        return true;
      }

      return !isRegimentOverviewSheet(
        name,
        regiment
      );
    });
}

async function handleCompanyAutocomplete(interaction) {
  const focused =
    interaction.options.getFocused(true);

  const isAddMemberCompany =
    focused.name === "company";

  const isTransferCompany =
    focused.name === "new_company";

  if (!isAddMemberCompany && !isTransferCompany) {
    await interaction.respond([]);
    return;
  }

  const regimentOptionName =
    isTransferCompany
      ? "new_regiment"
      : "regiment";

  const regimentValue =
    interaction.options.getString(
      regimentOptionName
    );

  if (!regimentValue) {
    await interaction.respond([
      {
        name: "Select a regiment first",
        value: "Select a regiment first"
      }
    ]);
    return;
  }

  const regiment = resolveRegiment(regimentValue);

  let companies =
    await getCompanySheetNames(
      regiment
    );

  if (
    interaction.commandName ===
    "attendance"
  ) {
    companies =
      companies.filter(
        company =>
          !isAttendanceExcludedCompany(
            company
          )
      );
  }

  const searchText = normalizeText(
    focused.value
  );

  const suggestions = companies
    .filter(company =>
      !searchText ||
      normalizeText(company).includes(searchText)
    )
    .slice(0, 25)
    .map(company => ({
      name: company.slice(0, 100),
      value: company.slice(0, 100)
    }));

  await interaction.respond(suggestions);
}



async function handleAttendanceDayAutocomplete(
  interaction
) {
  const focused =
    interaction.options.getFocused(true);

  if (focused.name !== "day") {
    await interaction.respond([]);
    return;
  }

  const company =
    interaction.options.getString(
      "company"
    );

  if (!company) {
    await interaction.respond([]);
    return;
  }

  if (isAttendanceExcludedCompany(company)) {
    await interaction.respond([]);
    return;
  }

  const availableDays =
    isSecondKrumperCompany(company)
      ? ["Saturday", "Sunday"]
      : ATTENDANCE_DAYS;

  const searchText =
    normalizeText(focused.value);

  const suggestions =
    availableDays
      .filter(day =>
        !searchText ||
        normalizeText(day).includes(
          searchText
        )
      )
      .map(day => ({
        name: day,
        value: day
      }));

  await interaction.respond(
    suggestions
  );
}



/*
|--------------------------------------------------------------------------
| Multi-member attendance sessions
|--------------------------------------------------------------------------
|
| /attendance now works as:
| 1. Regiment
| 2. Company
| 3. Day
| 4. Select up to 15 Discord members total
| 5. Choose one attendance status for the whole selected group
|--------------------------------------------------------------------------
*/

const ATTENDANCE_SESSION_TTL_MS =
  10 * 60 * 1000;

const ATTENDANCE_MAX_MEMBERS = 15;

const attendanceSessions =
  new Map();

function createAttendanceSessionToken(
  interaction
) {
  return (
    `${interaction.id}-${Date.now()}`
      .replace(/[^a-zA-Z0-9_-]/g, "")
      .slice(-70)
  );
}

function getAttendanceSession(
  token,
  interaction
) {
  const session =
    attendanceSessions.get(token);

  if (!session) {
    return null;
  }

  if (
    session.expiresAt < Date.now()
  ) {
    attendanceSessions.delete(token);
    return null;
  }

  if (
    session.ownerId !==
    interaction.user.id
  ) {
    return null;
  }

  if (
    session.guildId &&
    interaction.guildId &&
    session.guildId !==
      interaction.guildId
  ) {
    return null;
  }

  return session;
}

function scheduleAttendanceSessionExpiry(
  token
) {
  setTimeout(() => {
    const session =
      attendanceSessions.get(token);

    if (
      session &&
      session.expiresAt <= Date.now()
    ) {
      attendanceSessions.delete(token);
    }
  }, ATTENDANCE_SESSION_TTL_MS + 1000);
}

function buildAttendanceDayComponents(
  session
) {
  const availableDays =
    isSecondKrumperCompany(
      session.company
    )
      ? ["Saturday", "Sunday"]
      : ATTENDANCE_DAYS;

  const dayMenu =
    new StringSelectMenuBuilder()
      .setCustomId(
        `attendance_day:${session.token}`
      )
      .setPlaceholder(
        "Choose the attendance day"
      )
      .setMinValues(1)
      .setMaxValues(1)
      .addOptions(
        availableDays.map(day => ({
          label: day,
          value: day
        }))
      );

  const cancelButton =
    new ButtonBuilder()
      .setCustomId(
        `attendance_cancel:${session.token}`
      )
      .setLabel("Cancel")
      .setStyle(ButtonStyle.Danger);

  return [
    new ActionRowBuilder()
      .addComponents(dayMenu),
    new ActionRowBuilder()
      .addComponents(cancelButton)
  ];
}

function buildAttendanceMemberComponents(
  session
) {
  const remaining =
    Math.max(
      0,
      ATTENDANCE_MAX_MEMBERS -
      session.userIds.size
    );

  const rows = [];

  if (remaining > 0) {
    const userMenu =
      new UserSelectMenuBuilder()
        .setCustomId(
          `attendance_users:${session.token}`
        )
        .setPlaceholder(
          `Select members (${remaining} remaining)`
        )
        .setMinValues(1)
        .setMaxValues(
          Math.min(
            ATTENDANCE_MAX_MEMBERS,
            remaining
          )
        );

    rows.push(
      new ActionRowBuilder()
        .addComponents(userMenu)
    );
  }

  const continueButton =
    new ButtonBuilder()
      .setCustomId(
        `attendance_continue:${session.token}`
      )
      .setLabel(
        `Continue (${session.userIds.size})`
      )
      .setStyle(ButtonStyle.Success)
      .setDisabled(
        session.userIds.size === 0
      );

  const clearButton =
    new ButtonBuilder()
      .setCustomId(
        `attendance_clear:${session.token}`
      )
      .setLabel("Clear Members")
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(
        session.userIds.size === 0
      );

  const cancelButton =
    new ButtonBuilder()
      .setCustomId(
        `attendance_cancel:${session.token}`
      )
      .setLabel("Cancel")
      .setStyle(ButtonStyle.Danger);

  rows.push(
    new ActionRowBuilder()
      .addComponents(
        continueButton,
        clearButton,
        cancelButton
      )
  );

  return rows;
}

function buildAttendanceStatusComponents(
  session
) {
  const statusMenu =
    new StringSelectMenuBuilder()
      .setCustomId(
        `attendance_status:${session.token}`
      )
      .setPlaceholder(
        "Choose attendance for all selected members"
      )
      .setMinValues(1)
      .setMaxValues(1)
      .addOptions(
        ATTENDANCE_STATUS_CHOICES.map(
          status => ({
            label: status,
            value: status
          })
        )
      );

  const backButton =
    new ButtonBuilder()
      .setCustomId(
        `attendance_back:${session.token}`
      )
      .setLabel("Back to Members")
      .setStyle(ButtonStyle.Secondary);

  const cancelButton =
    new ButtonBuilder()
      .setCustomId(
        `attendance_cancel:${session.token}`
      )
      .setLabel("Cancel")
      .setStyle(ButtonStyle.Danger);

  return [
    new ActionRowBuilder()
      .addComponents(statusMenu),
    new ActionRowBuilder()
      .addComponents(
        backButton,
        cancelButton
      )
  ];
}

function attendanceSessionHeader(
  session
) {
  return [
    "**Attendance Entry**",
    "",
    `**Regiment:** ${session.regimentDisplayName}`,
    `**Company:** ${session.company}`,
    session.day
      ? `**Day:** ${session.day}`
      : "**Day:** Not selected",
    `**Members Selected:** ${session.userIds.size}`
  ];
}

async function handleAttendanceComponent(
  interaction
) {
  const customId =
    String(interaction.customId || "");

  if (
    !customId.startsWith(
      "attendance_"
    )
  ) {
    return false;
  }

  const separatorIndex =
    customId.indexOf(":");

  if (separatorIndex === -1) {
    return false;
  }

  const action =
    customId.slice(
      0,
      separatorIndex
    );

  const token =
    customId.slice(
      separatorIndex + 1
    );

  const session =
    getAttendanceSession(
      token,
      interaction
    );

  if (!session) {
    await interaction.reply({
      content:
        "This attendance session expired or belongs to another user. Run `/attendance` again.",
      flags:
        MessageFlags.Ephemeral
    });
    return true;
  }

  session.expiresAt =
    Date.now() +
    ATTENDANCE_SESSION_TTL_MS;

  if (
    action ===
    "attendance_cancel"
  ) {
    attendanceSessions.delete(token);

    await interaction.update({
      content:
        "Attendance entry cancelled.",
      components: []
    });

    return true;
  }

  if (
    action ===
    "attendance_day"
  ) {
    const day =
      interaction.values?.[0];

    try {
      getAttendanceColumn(
        session.company,
        day
      );
    } catch (error) {
      await interaction.reply({
        content:
          error?.message ===
          "SECOND_KRUMPER_WEEKEND_ONLY"
            ? "2. Krümper-Kompanie only allows Saturday or Sunday."
            : "That attendance day is not valid for this company.",
        flags:
          MessageFlags.Ephemeral
      });
      return true;
    }

    session.day = day;

    await interaction.update({
      content: [
        ...attendanceSessionHeader(
          session
        ),
        "",
        `Select up to **${ATTENDANCE_MAX_MEMBERS}** members total.`,
        "You may use the selector again until the 15-member limit is reached.",
        "When everyone is selected, click **Continue**."
      ].join("\n"),
      components:
        buildAttendanceMemberComponents(
          session
        )
    });

    return true;
  }

  if (
    action ===
    "attendance_users"
  ) {
    const beforeCount =
      session.userIds.size;

    for (
      const userId of
      interaction.values || []
    ) {
      if (
        session.userIds.size >=
        ATTENDANCE_MAX_MEMBERS
      ) {
        break;
      }

      session.userIds.add(userId);
    }

    const addedCount =
      session.userIds.size -
      beforeCount;

    const atLimit =
      session.userIds.size >=
      ATTENDANCE_MAX_MEMBERS;

    await interaction.update({
      content: [
        ...attendanceSessionHeader(
          session
        ),
        "",
        `Added **${addedCount}** member(s). **${session.userIds.size}/${ATTENDANCE_MAX_MEMBERS}** selected.`,
        atLimit
          ? "The 15-member limit has been reached. Click **Continue** to choose attendance."
          : "Select more members or click **Continue**."
      ].join("\n"),
      components:
        buildAttendanceMemberComponents(
          session
        )
    });

    return true;
  }

  if (
    action ===
    "attendance_clear"
  ) {
    session.userIds.clear();

    await interaction.update({
      content: [
        ...attendanceSessionHeader(
          session
        ),
        "",
        "The member selection was cleared.",
        "Select members again."
      ].join("\n"),
      components:
        buildAttendanceMemberComponents(
          session
        )
    });

    return true;
  }

  if (
    action ===
    "attendance_back"
  ) {
    await interaction.update({
      content: [
        ...attendanceSessionHeader(
          session
        ),
        "",
        "Add more members or continue when ready."
      ].join("\n"),
      components:
        buildAttendanceMemberComponents(
          session
        )
    });

    return true;
  }

  if (
    action ===
    "attendance_continue"
  ) {
    if (
      session.userIds.size === 0
    ) {
      await interaction.reply({
        content:
          "Select at least one member before continuing.",
        flags:
          MessageFlags.Ephemeral
      });
      return true;
    }

    if (
      session.userIds.size >
      ATTENDANCE_MAX_MEMBERS
    ) {
      await interaction.reply({
        content:
          `Attendance is limited to ${ATTENDANCE_MAX_MEMBERS} members at a time.`,
        flags:
          MessageFlags.Ephemeral
      });
      return true;
    }

    await interaction.update({
      content: [
        ...attendanceSessionHeader(
          session
        ),
        "",
        "Choose the attendance status that should be applied to **all selected members**."
      ].join("\n"),
      components:
        buildAttendanceStatusComponents(
          session
        )
    });

    return true;
  }

  if (
    action ===
    "attendance_status"
  ) {
    const attendance =
      String(
        interaction.values?.[0] || ""
      ).trim();

    if (
      !ATTENDANCE_STATUS_CHOICES.some(
        value =>
          normalizeText(value) ===
          normalizeText(attendance)
      )
    ) {
      await interaction.reply({
        content:
          "That attendance status is not configured.",
        flags:
          MessageFlags.Ephemeral
      });
      return true;
    }

    await interaction.deferUpdate();

    const updated = [];
    const skipped = [];
    const previousValues = [];

    for (
      const userId of
      session.userIds
    ) {
      try {
        const row =
          await findMemberRowInCompanyByDiscordId({
            spreadsheetId:
              session.spreadsheetId,
            sheetName:
              session.company,
            discordId:
              userId
          });

        if (!row) {
          skipped.push({
            userId,
            reason:
              "not assigned to the selected company"
          });
          continue;
        }

        const memberRecord =
          await getMemberRecord({
            spreadsheetId:
              session.spreadsheetId,
            sheetName:
              session.company,
            row
          });

        const attendanceResult =
          await setAttendanceMarker({
            spreadsheetId:
              session.spreadsheetId,
            sheetName:
              session.company,
            row,
            day:
              session.day,
            attendance
          });

        updated.push({
          userId,
          row,
          robloxUsername:
            memberRecord.robloxUsername,
          cell:
            attendanceResult.range
        });

        previousValues.push(
          attendanceResult.previousValue ||
          "Blank"
        );
      } catch (error) {
        console.error(
          `Attendance update failed for Discord user ${userId}:`
        );
        console.error(error);

        skipped.push({
          userId,
          reason:
            error?.message ||
            "unknown spreadsheet error"
        });
      }
    }

    attendanceSessions.delete(token);

    await sendOrbatLog({
      interaction,
      category: "Attendance",
      action:
        "Bulk Attendance Updated",
      affectedMember: null,
      robloxUsername:
        `${updated.length} member(s)`,
      changes: [
        {
          label: "Regiment",
          after:
            session.regimentDisplayName
        },
        {
          label: "Company",
          after:
            session.company
        },
        {
          label: "Day",
          after:
            session.day
        },
        {
          label: "Attendance",
          after:
            attendance
        },
        {
          label: "Members Updated",
          after:
            String(updated.length)
        },
        {
          label: "Members Skipped",
          after:
            String(skipped.length)
        }
      ],
      notes:
        updated.length > 0
          ? updated
              .slice(0, 20)
              .map(
                member =>
                  `<@${member.userId}> → ${member.cell}`
              )
              .join("\n")
          : "No members were updated."
    });

    const updatedPreview =
      updated
        .slice(0, 20)
        .map(
          member =>
            `• <@${member.userId}> — ${member.cell}`
        );

    const skippedPreview =
      skipped
        .slice(0, 15)
        .map(
          member =>
            `• <@${member.userId}> — ${member.reason}`
        );

    const responseLines = [
      "**Attendance Updated**",
      "",
      `**Regiment:** ${session.regimentDisplayName}`,
      `**Company:** ${session.company}`,
      `**Day:** ${session.day}`,
      `**Attendance:** ${attendance}`,
      `**Members Updated:** ${updated.length}`,
      `**Members Skipped:** ${skipped.length}`
    ];

    if (
      updatedPreview.length > 0
    ) {
      responseLines.push(
        "",
        "**Updated:**",
        ...updatedPreview
      );
    }

    if (
      updated.length >
      updatedPreview.length
    ) {
      responseLines.push(
        `• …and ${updated.length - updatedPreview.length} more`
      );
    }

    if (
      skippedPreview.length > 0
    ) {
      responseLines.push(
        "",
        "**Skipped:**",
        ...skippedPreview
      );
    }

    if (
      skipped.length >
      skippedPreview.length
    ) {
      responseLines.push(
        `• …and ${skipped.length - skippedPreview.length} more`
      );
    }

    await interaction.editReply({
      content:
        responseLines.join("\n"),
      components: []
    });

    return true;
  }

  return false;
}

/*
|--------------------------------------------------------------------------
| Command role permissions
|--------------------------------------------------------------------------
|
| A Discord member may use /addmember when they have at least one of the
| roles listed below. Role-name matching is case-insensitive.
|--------------------------------------------------------------------------
*/

const ALLOWED_COMMAND_ROLES = new Set([
  "unteroffizier",
  "kompagnieoffizier",
  "stabsoffizier",
  "generalstab",
  "general-major",
  "general-feldmarschall",
  "könig friedrich wilhelm iii"
]);

function memberHasCommandRole(interaction) {
  if (!interaction.inGuild()) {
    return false;
  }

  const member = interaction.member;

  if (!member || !member.roles) {
    return false;
  }

  /*
   * GuildMember instances provide a role cache. Raw API interaction members
   * may instead provide an array of role IDs, so the GuildMember path is
   * required here to check role names safely.
   */
  if (!member.roles.cache) {
    return false;
  }

  return member.roles.cache.some(role =>
    ALLOWED_COMMAND_ROLES.has(
      String(role.name || "")
        .trim()
        .toLowerCase()
    )
  );
}


/*
|--------------------------------------------------------------------------
| ORBAT audit logging
|--------------------------------------------------------------------------
*/

function safeLogValue(value) {
  const text = String(value ?? "").trim();
  return text || "Not set";
}

async function sendOrbatLog({
  interaction,
  category,
  action,
  affectedMember,
  robloxUsername,
  changes = [],
  notes = null
}) {
  try {
    if (!interaction.guild) {
      console.warn(
        "ORBAT audit log skipped because the command was not used in a server."
      );
      return false;
    }

    let logChannel =
      interaction.guild.channels.cache.find(
        channel =>
          channel.name === ORBAT_LOG_CHANNEL_NAME &&
          channel.isTextBased()
      );

    if (!logChannel) {
      const channels =
        await interaction.guild.channels.fetch();

      logChannel =
        channels.find(
          channel =>
            channel?.name === ORBAT_LOG_CHANNEL_NAME &&
            channel.isTextBased()
        );
    }

    if (!logChannel) {
      console.warn(
        `Create a text channel named #${ORBAT_LOG_CHANNEL_NAME} to receive ORBAT audit logs.`
      );
      return false;
    }

    const performedByName =
      interaction.member?.displayName ||
      interaction.user.globalName ||
      interaction.user.username;

    const changeText =
      changes.length > 0
        ? changes.map(change => {
            const label = safeLogValue(change.label);
            const hasBefore =
              change.before !== undefined;
            const hasAfter =
              change.after !== undefined;

            if (hasBefore && hasAfter) {
              return (
                `**${label}**\n` +
                `Before: ${safeLogValue(change.before)}\n` +
                `After: ${safeLogValue(change.after)}`
              );
            }

            if (hasAfter) {
              return (
                `**${label}:** ` +
                safeLogValue(change.after)
              );
            }

            return (
              `**${label}:** ` +
              safeLogValue(change.before)
            );
          }).join("\n\n")
        : "No detailed changes were supplied.";

    const affectedMemberText =
      affectedMember
        ? `<@${affectedMember.id}>\n${affectedMember.username} (${affectedMember.id})`
        : "Not set";

    const embed =
      new EmbedBuilder()
        .setTitle("Prussian ORBAT Audit Log")
        .setDescription(
          `**Category:** ${safeLogValue(category)}\n` +
          `**Action:** ${safeLogValue(action)}`
        )
        .addFields(
          {
            name: "Performed By",
            value:
              `<@${interaction.user.id}>\n` +
              `${performedByName} (${interaction.user.id})`,
            inline: true
          },
          {
            name: "Affected Member",
            value: affectedMemberText,
            inline: true
          },
          {
            name: "Roblox Username",
            value: safeLogValue(robloxUsername),
            inline: true
          },
          {
            name: "Changes",
            value: changeText.slice(0, 1024),
            inline: false
          }
        )
        .setFooter({
          text:
            `Command: /${interaction.commandName} • ` +
            `Executed by ${interaction.user.username}`
        })
        .setTimestamp();

    if (notes) {
      embed.addFields({
        name: "Notes",
        value: safeLogValue(notes).slice(0, 1024),
        inline: false
      });
    }

    await logChannel.send({
      embeds: [embed]
    });

    return true;
  } catch (error) {
    console.error("Failed to send ORBAT audit log:");
    console.error(error);
    return false;
  }
}

/*
|--------------------------------------------------------------------------
| Bot ready event
|--------------------------------------------------------------------------
*/

client.once(Events.ClientReady, readyClient => {
  console.log(`Logged in as ${readyClient.user.tag}`);
  console.log("Configured regiment spreadsheets:");

  for (const regiment of REGIMENTS) {
    console.log(
      `- ${regiment.displayName}: ${regiment.spreadsheetId}`
    );
  }

  console.log("Prussian ORBAT bot is online.");
});

/*
|--------------------------------------------------------------------------
| Slash-command handler
|--------------------------------------------------------------------------
*/

client.on(
  Events.InteractionCreate,
  async interaction => {
    if (interaction.isAutocomplete()) {
      const isCompanyAutocompleteCommand =
        interaction.commandName ===
          "addmember" ||
        interaction.commandName ===
          "transfer" ||
        interaction.commandName ===
          "attendance" ||
        interaction.commandName ===
          "orginize" ||
        interaction.commandName ===
          "companyinfo" ||
        interaction.commandName ===
          "missingattendance" ||
        interaction.commandName ===
          "roster" ||
        interaction.commandName ===
          "attendanceview";

      if (!isCompanyAutocompleteCommand) {
        return;
      }

      try {
        await handleCompanyAutocomplete(
          interaction
        );
      } catch (error) {
        console.error(
          "Autocomplete failed:"
        );
        console.error(error);

        try {
          await interaction.respond([]);
        } catch {
          // Ignore expired or already-answered autocomplete requests.
        }
      }

      return;
    }

    if (
      interaction.isUserSelectMenu() ||
      interaction.isStringSelectMenu() ||
      interaction.isButton()
    ) {
      try {
        const handled =
          await handleAttendanceComponent(
            interaction
          );

        if (handled) {
          return;
        }
      } catch (error) {
        console.error(
          "Attendance component failed:"
        );
        console.error(error);

        try {
          if (
            interaction.deferred ||
            interaction.replied
          ) {
            await interaction.editReply({
              content:
                "The attendance session encountered an error.",
              components: []
            });
          } else {
            await interaction.reply({
              content:
                "The attendance session encountered an error.",
              flags:
                MessageFlags.Ephemeral
            });
          }
        } catch {
          // Ignore response failures for expired component interactions.
        }

        return;
      }
    }

    if (!interaction.isChatInputCommand()) {
      return;
    }

    if (
      interaction.commandName !== "addmember" &&
      interaction.commandName !== "removemember" &&
      interaction.commandName !== "rank" &&
      interaction.commandName !== "transfer" &&
      interaction.commandName !== "attendance" &&
      interaction.commandName !== "orginize" &&
      interaction.commandName !== "memberinfo" &&
      interaction.commandName !== "companyinfo" &&
      interaction.commandName !== "missingattendance" &&
      interaction.commandName !== "getsheet" &&
      interaction.commandName !== "roster" &&
      interaction.commandName !== "strength" &&
      interaction.commandName !== "attendanceview" &&
      interaction.commandName !== "audit"
    ) {
      return;
    }

    if (!memberHasCommandRole(interaction)) {
      await interaction.reply({
        content:
          "You do not have permission to use this command. " +
          "You must have one of these roles: Unteroffizier, " +
          "Kompagnieoffizier, Stabsoffizier, Generalstab, " +
          "General-Major, General-Feldmarschall, or König Friedrich Wilhelm III.",
        flags: MessageFlags.Ephemeral
      });
      return;
    }


    if (interaction.commandName === "roster") {
      try {
        await interaction.deferReply({
          flags: MessageFlags.Ephemeral
        });

        const regimentValue =
          interaction.options
            .getString(
              "regiment",
              true
            )
            .trim();

        const company =
          interaction.options
            .getString(
              "company",
              true
            )
            .trim();

        const regiment =
          resolveRegiment(
            regimentValue
          );

        const availableCompanies =
          await getCompanySheetNames(
            regiment
          );

        const matchedCompany =
          availableCompanies.find(
            name =>
              normalizeText(name) ===
              normalizeText(company)
          );

        if (!matchedCompany) {
          await interaction.editReply(
            "The selected company could not be found."
          );
          return;
        }

        const summary =
          await getCompanyRosterSummary({
            spreadsheetId:
              regiment.spreadsheetId,
            sheetName:
              matchedCompany
          });

        const members =
          [...summary.members]
            .sort((a, b) => {
              const rankA =
                RANK_SORT_PRIORITY.get(
                  normalizeText(a.rank)
                ) ?? 999;

              const rankB =
                RANK_SORT_PRIORITY.get(
                  normalizeText(b.rank)
                ) ?? 999;

              if (rankA !== rankB) {
                return rankA - rankB;
              }

              return a.robloxUsername.localeCompare(
                b.robloxUsername,
                undefined,
                {
                  sensitivity: "base"
                }
              );
            });

        const rosterLines =
          members
            .slice(0, 30)
            .map((member, index) => {
              const identity =
                member.discordId
                  ? `<@${member.discordId}>`
                  : member.robloxUsername ||
                    "Unknown";

              return (
                `${index + 1}. **${member.rank || "No Rank"}** — ` +
                `${identity}` +
                (
                  member.robloxUsername
                    ? ` (${member.robloxUsername})`
                    : ""
                )
              );
            });

        const lines = [
          "**Kompanie Roster**",
          "",
          `**Regiment:** ${regiment.displayName}`,
          `**Kompanie:** ${matchedCompany}`,
          `**Strength:** ${members.length}`,
          "",
          ...(
            rosterLines.length
              ? rosterLines
              : ["No members found."]
          )
        ];

        if (
          members.length >
          rosterLines.length
        ) {
          lines.push(
            `…and ${members.length - rosterLines.length} more member(s).`
          );
        }

        await interaction.editReply(
          lines.join("\n").slice(0, 1950)
        );

        return;
      } catch (error) {
        console.error(
          "Failed to retrieve roster:"
        );
        console.error(error);

        await interaction.editReply(
          `Roster could not be retrieved: ${error?.message || "Unknown error."}`
        );

        return;
      }
    }

    if (interaction.commandName === "strength") {
      try {
        await interaction.deferReply({
          flags: MessageFlags.Ephemeral
        });

        const regimentValue =
          interaction.options.getString(
            "regiment",
            false
          );

        const targetRegiments =
          regimentValue
            ? [
                resolveRegiment(
                  regimentValue
                )
              ]
            : REGIMENTS;

        const regimentResults = [];
        let grandTotal = 0;

        for (
          const regiment of
          targetRegiments
        ) {
          const companies =
            await getCompanySheetNames(
              regiment
            );

          const companyResults = [];
          let regimentTotal = 0;

          for (
            const company of companies
          ) {
            const summary =
              await getCompanyRosterSummary({
                spreadsheetId:
                  regiment.spreadsheetId,
                sheetName:
                  company
              });

            companyResults.push({
              company,
              strength:
                summary.members.length
            });

            regimentTotal +=
              summary.members.length;
          }

          regimentResults.push({
            regiment,
            regimentTotal,
            companyResults
          });

          grandTotal +=
            regimentTotal;
        }

        const lines = [
          regimentValue
            ? "**Regiment Strength**"
            : "**Prussian Army Strength**",
          "",
          `**Total Strength:** ${grandTotal}`
        ];

        for (
          const result of
          regimentResults
        ) {
          lines.push(
            "",
            `**${result.regiment.displayName}: ${result.regimentTotal}**`
          );

          if (regimentValue) {
            for (
              const companyResult of
              result.companyResults
            ) {
              lines.push(
                `• ${companyResult.company}: ${companyResult.strength}`
              );
            }
          }
        }

        await interaction.editReply(
          lines.join("\n").slice(0, 1950)
        );

        return;
      } catch (error) {
        console.error(
          "Failed to retrieve strength:"
        );
        console.error(error);

        await interaction.editReply(
          `Strength could not be retrieved: ${error?.message || "Unknown error."}`
        );

        return;
      }
    }

    if (
      interaction.commandName ===
      "attendanceview"
    ) {
      try {
        await interaction.deferReply({
          flags: MessageFlags.Ephemeral
        });

        const regimentValue =
          interaction.options
            .getString(
              "regiment",
              true
            )
            .trim();

        const company =
          interaction.options
            .getString(
              "company",
              true
            )
            .trim();

        const day =
          interaction.options
            .getString(
              "day",
              true
            )
            .trim();

        const regiment =
          resolveRegiment(
            regimentValue
          );

        const companies =
          await getCompanySheetNames(
            regiment
          );

        const matchedCompany =
          companies.find(
            name =>
              normalizeText(name) ===
              normalizeText(company)
          );

        if (!matchedCompany) {
          await interaction.editReply(
            "The selected company could not be found."
          );
          return;
        }

        if (
          isAttendanceExcludedCompany(
            matchedCompany
          )
        ) {
          await interaction.editReply(
            "Attendance is not tracked for that company."
          );
          return;
        }

        let view;

        try {
          view =
            await getCompanyAttendanceView({
              spreadsheetId:
                regiment.spreadsheetId,
              sheetName:
                matchedCompany,
              day
            });
        } catch (
          attendanceError
        ) {
          if (
            attendanceError?.message ===
            "SECOND_KRUMPER_WEEKEND_ONLY"
          ) {
            await interaction.editReply(
              "2. Krümper-Kompanie only tracks Saturday and Sunday attendance."
            );
            return;
          }

          throw attendanceError;
        }

        const lines = [
          "**Attendance View**",
          "",
          `**Regiment:** ${regiment.displayName}`,
          `**Kompanie:** ${matchedCompany}`,
          `**Day:** ${day}`,
          `**Total Members:** ${view.members.length}`,
          ""
        ];

        for (
          const status of
          [
            ...ATTENDANCE_STATUS_CHOICES,
            "BLANK"
          ]
        ) {
          const members =
            view.groups.get(status) || [];

          if (!members.length) {
            continue;
          }

          lines.push(
            `**${status}: ${members.length}**`
          );

          for (
            const member of
            members.slice(0, 8)
          ) {
            const identity =
              member.discordId
                ? `<@${member.discordId}>`
                : member.robloxUsername ||
                  "Unknown";

            lines.push(
              `• ${identity}`
            );
          }

          if (members.length > 8) {
            lines.push(
              `• …and ${members.length - 8} more`
            );
          }

          lines.push("");
        }

        await interaction.editReply(
          lines.join("\n").slice(0, 1950)
        );

        return;
      } catch (error) {
        console.error(
          "Failed to retrieve attendance view:"
        );
        console.error(error);

        await interaction.editReply(
          `Attendance view could not be retrieved: ${error?.message || "Unknown error."}`
        );

        return;
      }
    }

    if (interaction.commandName === "audit") {
      try {
        await interaction.deferReply({
          flags: MessageFlags.Ephemeral
        });

        const regimentValue =
          interaction.options.getString(
            "regiment",
            false
          );

        const targetRegiments =
          regimentValue
            ? [
                resolveRegiment(
                  regimentValue
                )
              ]
            : REGIMENTS;

        const issues =
          await auditOrbatRegiments(
            targetRegiments
          );

        const lines = [
          "**Grand ORBAT Audit**",
          "",
          `**Scope:** ${
            regimentValue
              ? targetRegiments[0].displayName
              : "All Regiments"
          }`,
          `**Issues Found:** ${issues.length}`,
          ""
        ];

        if (!issues.length) {
          lines.push(
            "No roster issues were found."
          );
        } else {
          lines.push(
            "**Issues:**"
          );

          for (
            const issue of
            issues.slice(0, 18)
          ) {
            lines.push(
              `• ${issue}`
            );
          }

          if (issues.length > 18) {
            lines.push(
              `• …and ${issues.length - 18} more issue(s).`
            );
          }
        }

        await interaction.editReply(
          lines.join("\n").slice(0, 1950)
        );

        return;
      } catch (error) {
        console.error(
          "Failed to audit ORBAT:"
        );
        console.error(error);

        await interaction.editReply(
          `The ORBAT audit could not be completed: ${error?.message || "Unknown error."}`
        );

        return;
      }
    }

    if (interaction.commandName === "getsheet") {
      try {
        await interaction.deferReply({
          flags:
            MessageFlags.Ephemeral
        });

        const regimentValue =
          interaction.options
            .getString(
              "regiment",
              true
            )
            .trim();

        const regiment =
          resolveRegiment(
            regimentValue
          );

        const sheetUrl =
          `https://docs.google.com/spreadsheets/d/${regiment.spreadsheetId}/edit`;

        await interaction.editReply(
          [
            "**Regiment ORBAT Sheet**",
            "",
            `**Regiment:** ${regiment.displayName}`,
            "",
            `[Open Google Sheet](${sheetUrl})`
          ].join("\n")
        );

        return;
      } catch (error) {
        console.error(
          "Failed to retrieve regiment sheet:"
        );
        console.error(error);

        const errorMessage =
          error?.message ||
          "Unknown sheet lookup error.";

        try {
          if (
            interaction.deferred ||
            interaction.replied
          ) {
            await interaction.editReply(
              `The regiment sheet could not be retrieved: ${errorMessage}`
            );
          } else {
            await interaction.reply({
              content:
                `The regiment sheet could not be retrieved: ${errorMessage}`,
              flags:
                MessageFlags.Ephemeral
            });
          }
        } catch (replyError) {
          console.error(
            "Failed to send getsheet error reply:"
          );
          console.error(replyError);
        }

        return;
      }
    }

    if (interaction.commandName === "memberinfo") {
      try {
        await interaction.deferReply({ flags: MessageFlags.Ephemeral });

        const discordMember =
          interaction.options.getUser("discord_member", true);

        const existingMember =
          await findMemberByDiscordId(discordMember.id);

        if (!existingMember) {
          await interaction.editReply(
            [
              "That member was not found in the Grand ORBAT.",
              "",
              `**Discord Member:** ${discordMember}`,
              `**Discord ID:** ${discordMember.id}`
            ].join("\n")
          );
          return;
        }

        const memberRecord = await getMemberRecord({
          spreadsheetId: existingMember.regiment.spreadsheetId,
          sheetName: existingMember.companyName,
          row: existingMember.row
        });

        const attendanceSummary = await getMemberAttendanceSummary({
          spreadsheetId: existingMember.regiment.spreadsheetId,
          sheetName: existingMember.companyName,
          row: existingMember.row
        });

        await interaction.editReply(
          [
            "**Grand ORBAT Member Information**",
            "",
            `**Discord Member:** ${discordMember}`,
            `**Roblox Username:** ${memberRecord.robloxUsername || "Not set"}`,
            `**Regiment:** ${existingMember.regiment.displayName}`,
            `**Kompanie:** ${existingMember.companyName}`,
            `**Rank:** ${memberRecord.rank || "Not set"}`,
            `**Timezone:** ${memberRecord.timezone || "Not set"}`,
            `**Sheet Row:** ${existingMember.row}`,
            "",
            "**Current Attendance**",
            ...attendanceSummary.lines
          ].join("\n")
        );
        return;
      } catch (error) {
        console.error("Failed to retrieve member information:", error);
        await interaction.editReply(
          `Member information could not be retrieved: ${error?.message || "Unknown error."}`
        );
        return;
      }
    }

    if (interaction.commandName === "companyinfo") {
      try {
        await interaction.deferReply({ flags: MessageFlags.Ephemeral });

        const regimentValue =
          interaction.options.getString("regiment", true).trim();
        const company =
          interaction.options.getString("company", true).trim();

        const regiment = resolveRegiment(regimentValue);
        const availableCompanies = await getCompanySheetNames(regiment);
        const matchedCompany = availableCompanies.find(
          name => normalizeText(name) === normalizeText(company)
        );

        if (!matchedCompany) {
          await interaction.editReply("The selected company could not be found.");
          return;
        }

        const summary = await getCompanyRosterSummary({
          spreadsheetId: regiment.spreadsheetId,
          sheetName: matchedCompany
        });

        const maxSlots =
          getLastMemberRow(matchedCompany) - FIRST_MEMBER_ROW + 1;

        const rankLines =
          summary.rankBreakdown.length
            ? summary.rankBreakdown.map(([rank, count]) => `• **${rank}:** ${count}`)
            : ["• No members found."];

        await interaction.editReply(
          [
            "**Kompanie Information**",
            "",
            `**Regiment:** ${regiment.displayName}`,
            `**Kompanie:** ${matchedCompany}`,
            `**Current Strength:** ${summary.members.length}`,
            `**Roster Capacity:** ${maxSlots}`,
            `**Open Slots:** ${Math.max(0, maxSlots - summary.members.length)}`,
            "",
            "**Rank Breakdown**",
            ...rankLines
          ].join("\n")
        );
        return;
      } catch (error) {
        console.error("Failed to retrieve company information:", error);
        await interaction.editReply(
          `Company information could not be retrieved: ${error?.message || "Unknown error."}`
        );
        return;
      }
    }

    if (interaction.commandName === "missingattendance") {
      try {
        await interaction.deferReply({ flags: MessageFlags.Ephemeral });

        const regimentValue =
          interaction.options.getString("regiment", true).trim();
        const company =
          interaction.options.getString("company", true).trim();
        const day =
          interaction.options.getString("day", true).trim();

        const regiment = resolveRegiment(regimentValue);
        const availableCompanies = await getCompanySheetNames(regiment);
        const matchedCompany = availableCompanies.find(
          name => normalizeText(name) === normalizeText(company)
        );

        if (!matchedCompany) {
          await interaction.editReply("The selected company could not be found.");
          return;
        }

        if (isAttendanceExcludedCompany(matchedCompany)) {
          await interaction.editReply(
            "Attendance is not tracked for that company."
          );
          return;
        }

        try {
          getAttendanceColumn(matchedCompany, day);
        } catch (attendanceError) {
          if (attendanceError?.message === "SECOND_KRUMPER_WEEKEND_ONLY") {
            await interaction.editReply(
              "2. Krümper-Kompanie only tracks Saturday and Sunday attendance."
            );
            return;
          }
          throw attendanceError;
        }

        const missingMembers = await getMissingAttendanceMembers({
          spreadsheetId: regiment.spreadsheetId,
          sheetName: matchedCompany,
          day
        });

        const memberLines = missingMembers.slice(0, 30).map(member => {
          const identity = member.discordId
            ? `<@${member.discordId}>`
            : member.robloxUsername || "Unknown member";
          return `• ${identity}${member.rank ? ` — ${member.rank}` : ""} — Row ${member.row}`;
        });

        const lines = [
          "**Missing Attendance**",
          "",
          `**Regiment:** ${regiment.displayName}`,
          `**Kompanie:** ${matchedCompany}`,
          `**Day:** ${day}`,
          `**Missing:** ${missingMembers.length}`,
          ""
        ];

        if (!missingMembers.length) {
          lines.push(
            "Everyone on the roster has an attendance marker for that day."
          );
        } else {
          lines.push("**Members Without Attendance:**", ...memberLines);
          if (missingMembers.length > memberLines.length) {
            lines.push(
              `• …and ${missingMembers.length - memberLines.length} more`
            );
          }
        }

        await interaction.editReply(lines.join("\n"));
        return;
      } catch (error) {
        console.error("Failed to retrieve missing attendance:", error);
        await interaction.editReply(
          `Missing attendance could not be retrieved: ${error?.message || "Unknown error."}`
        );
        return;
      }
    }

    if (interaction.commandName === "orginize") {
      try {
        await interaction.deferReply({
          flags:
            MessageFlags.Ephemeral
        });

        const regimentValue =
          interaction.options
            .getString(
              "regiment",
              true
            )
            .trim();

        const company =
          interaction.options
            .getString(
              "company",
              true
            )
            .trim();

        const regiment =
          resolveRegiment(
            regimentValue
          );

        const availableCompanies =
          await getCompanySheetNames(
            regiment
          );

        const matchedCompany =
          availableCompanies.find(
            availableCompany =>
              normalizeText(
                availableCompany
              ) ===
              normalizeText(company)
          );

        if (!matchedCompany) {
          await interaction.editReply(
            [
              "The selected company could not be found.",
              "",
              `**Regiment:** ${regiment.displayName}`,
              `**Company Submitted:** ${company}`,
              "",
              "Select an actual company from autocomplete and try again."
            ].join("\n")
          );
          return;
        }

        const memberCount =
          await sortCompanyByRank({
            spreadsheetId:
              regiment.spreadsheetId,
            sheetName:
              matchedCompany
          });

        await sendOrbatLog({
          interaction,
          category:
            "Organization",
          action:
            "Company Organized by Rank",
          affectedMember:
            null,
          robloxUsername:
            `${memberCount} member(s)`,
          changes: [
            {
              label: "Regiment",
              after:
                regiment.displayName
            },
            {
              label: "Company",
              after:
                matchedCompany
            },
            {
              label:
                "Members Organized",
              after:
                String(memberCount)
            }
          ],
          notes:
            "Company roster was sorted from highest rank to lowest rank. Attendance stayed attached to each member."
        });

        await interaction.editReply(
          [
            "**Company Organized**",
            "",
            `**Regiment:** ${regiment.displayName}`,
            `**Company:** ${matchedCompany}`,
            `**Members Organized:** ${memberCount}`,
            "",
            "The company has been sorted from highest rank to lowest rank.",
            "Attendance and other member-owned row data stayed with each member."
          ].join("\n")
        );

        return;
      } catch (error) {
        console.error(
          "Failed to organize company:"
        );
        console.error(error);

        const errorMessage =
          error?.message ||
          "Unknown organization error.";

        try {
          if (
            interaction.deferred ||
            interaction.replied
          ) {
            await interaction.editReply(
              `The company could not be organized: ${errorMessage}`
            );
          } else {
            await interaction.reply({
              content:
                `The company could not be organized: ${errorMessage}`,
              flags:
                MessageFlags.Ephemeral
            });
          }
        } catch (replyError) {
          console.error(
            "Failed to send organize error reply:"
          );
          console.error(replyError);
        }

        return;
      }
    }

    if (interaction.commandName === "attendance") {
      try {
        await interaction.deferReply({
          flags:
            MessageFlags.Ephemeral
        });

        const regimentValue =
          interaction.options
            .getString(
              "regiment",
              true
            )
            .trim();

        const company =
          interaction.options
            .getString(
              "company",
              true
            )
            .trim();

        const regiment =
          resolveRegiment(
            regimentValue
          );

        const availableCompanies =
          await getCompanySheetNames(
            regiment
          );

        const matchedCompany =
          availableCompanies.find(
            availableCompany =>
              normalizeText(
                availableCompany
              ) ===
              normalizeText(company)
          );

        if (!matchedCompany) {
          await interaction.editReply(
            [
              "The selected company could not be found.",
              "",
              `**Regiment:** ${regiment.displayName}`,
              `**Company Submitted:** ${company}`,
              "",
              "Select the company from autocomplete and try again."
            ].join("\n")
          );
          return;
        }

        if (
          isAttendanceExcludedCompany(
            matchedCompany
          )
        ) {
          await interaction.editReply(
            [
              "Attendance cannot be recorded for that sheet.",
              "",
              `**Company:** ${matchedCompany}`,
              "",
              "Attendance is disabled for Generalstab/command sheets, Garnison Kompanie, and 1. Krümper-Kompanie."
            ].join("\n")
          );
          return;
        }

        const token =
          createAttendanceSessionToken(
            interaction
          );

        const session = {
          token,
          ownerId:
            interaction.user.id,
          guildId:
            interaction.guildId,
          spreadsheetId:
            regiment.spreadsheetId,
          regimentKey:
            regiment.key,
          regimentDisplayName:
            regiment.displayName,
          company:
            matchedCompany,
          day: null,
          userIds:
            new Set(),
          expiresAt:
            Date.now() +
            ATTENDANCE_SESSION_TTL_MS
        };

        attendanceSessions.set(
          token,
          session
        );

        scheduleAttendanceSessionExpiry(
          token
        );

        await interaction.editReply({
          content: [
            ...attendanceSessionHeader(
              session
            ),
            "",
            "First choose the attendance day.",
            isSecondKrumperCompany(
              matchedCompany
            )
              ? "2. Krümper-Kompanie only allows Saturday or Sunday."
              : "Normal companies allow Monday through Sunday."
          ].join("\n"),
          components:
            buildAttendanceDayComponents(
              session
            )
        });

        return;
      } catch (error) {
        console.error(
          "Failed to start attendance session:"
        );
        console.error(error);

        const errorMessage =
          error?.message ||
          "Unknown attendance error.";

        try {
          if (
            interaction.deferred ||
            interaction.replied
          ) {
            await interaction.editReply({
              content:
                `Attendance could not be started: ${errorMessage}`,
              components: []
            });
          } else {
            await interaction.reply({
              content:
                `Attendance could not be started: ${errorMessage}`,
              flags:
                MessageFlags.Ephemeral
            });
          }
        } catch (replyError) {
          console.error(
            "Failed to send attendance error reply:"
          );
          console.error(replyError);
        }

        return;
      }
    }

    if (interaction.commandName === "transfer") {
      try {
        await interaction.deferReply({
          flags: MessageFlags.Ephemeral
        });

        const discordMember =
          interaction.options.getUser(
            "discord_member",
            true
          );

        const newRegimentValue =
          interaction.options
            .getString(
              "new_regiment",
              true
            )
            .trim();

        const newCompany =
          interaction.options
            .getString(
              "new_company",
              true
            )
            .trim();

        const requestedNewRank =
          interaction.options
            .getString(
              "new_rank",
              false
            )
            ?.trim() || null;

        const newPositionValue =
          interaction.options
            .getString(
              "new_position",
              false
            )
            ?.trim() || null;

        const existingMember =
          await findMemberByDiscordId(
            discordMember.id
          );

        if (!existingMember) {
          await interaction.editReply(
            [
              "That Discord member was not found in the Grand ORBAT.",
              "",
              `**Discord Member:** ${discordMember}`,
              `**Discord ID:** ${discordMember.id}`,
              "",
              "No transfer was made."
            ].join("\n")
          );
          return;
        }

        const newRegiment =
          resolveRegiment(
            newRegimentValue
          );

        let destinationPosition =
          null;

        if (
          isSchuetzenRegiment(
            newRegiment
          )
        ) {
          if (!newPositionValue) {
            await interaction.editReply(
              [
                "The transfer was not made.",
                "",
                "**Schlesisches Schützen-Bataillon uses a position system.**",
                "Choose **Company Commander**, **1. Platoon**, or **2. Platoon** in the `new_position` option and try again."
              ].join("\n")
            );
            return;
          }

          try {
            destinationPosition =
              resolveSchuetzenPosition(
                newPositionValue
              );
          } catch {
            await interaction.editReply(
              "That destination Schützen position is not configured."
            );
            return;
          }
        }

        const sameRegiment =
          existingMember.regiment.spreadsheetId ===
          newRegiment.spreadsheetId;

        const currentPosition =
          isSchuetzenRegiment(
            existingMember.regiment
          )
            ? getSchuetzenPositionForRow(
                existingMember.row
              )
            : null;

        const sameCompany =
          normalizeText(
            existingMember.companyName
          ) === normalizeText(newCompany) &&
          (
            !isSchuetzenRegiment(
              newRegiment
            ) ||
            currentPosition?.key ===
              destinationPosition?.key
          );

        const availableCompanies =
          await getCompanySheetNames(
            newRegiment
          );

        const matchedCompany =
          availableCompanies.find(
            company =>
              normalizeText(company) ===
              normalizeText(newCompany)
          );

        if (!matchedCompany) {
          await interaction.editReply(
            [
              "The selected destination company could not be found.",
              "",
              `**Regiment:** ${newRegiment.displayName}`,
              `**Company Submitted:** ${newCompany}`,
              "",
              "Select the company from autocomplete and try again."
            ].join("\n")
          );
          return;
        }

        const memberRecord =
          await getMemberRecord({
            spreadsheetId:
              existingMember.regiment.spreadsheetId,
            sheetName:
              existingMember.companyName,
            row:
              existingMember.row
          });

        if (!memberRecord.discordId) {
          memberRecord.discordId =
            discordMember.id;
        }

        const previousRank =
          String(memberRecord.rank || "").trim();

        const finalRank =
          requestedNewRank || previousRank;

        try {
          if (
            !isSchuetzenRegiment(
              newRegiment
            )
          ) {
            assertCompanyRankAllowed(
              matchedCompany,
              finalRank
            );
          }
        } catch (companyRankError) {
          if (
            companyRankError?.message ===
            "FIRST_KRUMPER_REKRUT_ONLY"
          ) {
            await interaction.editReply(
              "1. Krümper-Kompanie is recruit-only. Only members with the rank Rekrut may be transferred there."
            );
            return;
          }

          if (
            companyRankError?.message ===
            "REKRUT_FIRST_KRUMPER_ONLY"
          ) {
            await interaction.editReply(
              "A Rekrut cannot be transferred to that company. Rekruten may only be assigned to 1. Krümper-Kompanie."
            );
            return;
          }

          throw companyRankError;
        }

        const rankChanged =
          Boolean(requestedNewRank) &&
          normalizeText(requestedNewRank) !==
            normalizeText(previousRank);

        /*
         * When the selected regiment and company are unchanged, /transfer can
         * still be used to apply the optional new rank without moving rows.
         */
        if (sameRegiment && sameCompany) {
          if (!rankChanged) {
            await interaction.editReply(
              [
                "That member is already assigned to the selected regiment and company.",
                "",
                `**Regiment:** ${newRegiment.displayName}`,
                `**Company:** ${matchedCompany}`,
                `**Rank:** ${previousRank || "Not set"}`,
                "",
                requestedNewRank
                  ? "That member already has the selected rank."
                  : "No new rank was supplied.",
                "No spreadsheet cells were changed."
              ].join("\n")
            );
            return;
          }

          await updateMemberRank({
            spreadsheetId:
              existingMember.regiment.spreadsheetId,
            sheetName:
              existingMember.companyName,
            row:
              existingMember.row,
            rank:
              finalRank
          });

          await sortCompanyByRank({
            spreadsheetId:
              existingMember.regiment.spreadsheetId,
            sheetName:
              existingMember.companyName,
            platoon:
              currentPosition?.type ===
              "platoon"
                ? currentPosition
                : null
          });

          existingMember.row =
            await findMemberRowInCompanyByDiscordId({
              spreadsheetId:
                existingMember.regiment.spreadsheetId,
              sheetName:
                existingMember.companyName,
              discordId:
                discordMember.id
            }) || existingMember.row;

          let updatedNickname = null;
          let nicknameWarning = null;

          try {
            updatedNickname =
              await updateDiscordNickname({
                interaction,
                discordUserId:
                  discordMember.id,
                regiment:
                  existingMember.regiment,
                rank:
                  finalRank,
                robloxUsername:
                  memberRecord.robloxUsername
              });
          } catch (nicknameError) {
            console.error(
              "Transfer rank update succeeded, but nickname updating failed:"
            );
            console.error(nicknameError);

            nicknameWarning =
              nicknameError?.message ||
              "The Discord nickname could not be updated.";
          }

          const rankOnlyReply = [
            "Member rank updated successfully through /transfer.",
            "",
            `**Discord Member:** ${discordMember}`,
            `**Discord ID:** ${discordMember.id}`,
            `**Roblox Username:** ${memberRecord.robloxUsername || "Not set"}`,
            `**Regiment:** ${existingMember.regiment.displayName}`,
            `**Company:** ${existingMember.companyName}`,
            `**Spreadsheet Row:** ${existingMember.row}`,
            `**Previous Rank:** ${previousRank || "Not set"}`,
            `**New Rank:** ${finalRank}`
          ];

          if (updatedNickname) {
            rankOnlyReply.push(
              `**Discord Nickname:** ${updatedNickname}`
            );
          }

          if (nicknameWarning) {
            rankOnlyReply.push(
              "",
              "⚠️ The rank was updated, but the Discord nickname could not be updated.",
              `**Nickname Error:** ${nicknameWarning}`,
              "Make sure the bot has Manage Nicknames and its role is above the member's highest role."
            );
          }

          await sendOrbatLog({
            interaction,
            category: "Rank Management",
            action: "Rank Changed Through Transfer",
            affectedMember: discordMember,
            robloxUsername:
              memberRecord.robloxUsername,
            changes: [
              {
                label: "Regiment",
                after:
                  existingMember.regiment.displayName
              },
              {
                label: "Company",
                after:
                  existingMember.companyName
              },
              {
                label: "Rank",
                before: previousRank,
                after: finalRank
              },
              {
                label: "Spreadsheet Row",
                after: existingMember.row
              }
            ],
            notes: nicknameWarning
              ? "Rank updated, but the Discord nickname could not be updated."
              : null
          });

          await interaction.editReply(
            rankOnlyReply.join("\n")
          );
          return;
        }

        /*
         * Cross-company / cross-regiment transfers intentionally copy
         * ONLY the member data represented by C:F. Attendance markers
         * are company-specific and are NOT copied to the destination.
         */
        let destinationRow =
          await addMemberToSheet({
            spreadsheetId:
              newRegiment.spreadsheetId,
            sheetName:
              matchedCompany,
            robloxUsername:
              memberRecord.robloxUsername,
            discordId:
              memberRecord.discordId,
            rank:
              finalRank,
            timezone:
              memberRecord.timezone,
            position:
              destinationPosition
          });

        try {
          await removeMemberFromSheet({
            spreadsheetId:
              existingMember.regiment.spreadsheetId,
            sheetName:
              existingMember.companyName,
            row:
              existingMember.row
          });
        } catch (sourceClearError) {
          try {
            await removeMemberFromSheet({
              spreadsheetId:
                newRegiment.spreadsheetId,
              sheetName:
                matchedCompany,
              row:
                destinationRow
            });
          } catch (rollbackError) {
            console.error(
              "Transfer rollback also failed:"
            );
            console.error(rollbackError);
          }

          throw sourceClearError;
        }

        let timezoneWarning = null;

        if (memberRecord.timezone) {
          try {
            await processTimezoneWithAppsScript({
              spreadsheetId:
                newRegiment.spreadsheetId,
              sheetName:
                matchedCompany,
              row:
                destinationRow,
              timezone:
                memberRecord.timezone
            });
          } catch (webhookError) {
            console.error(
              "Transfer completed, but timezone processing failed:"
            );
            console.error(webhookError);

            timezoneWarning =
              webhookError?.message ||
              "The Apps Script webhook failed.";
          }
        }

        await sortCompanyByRank({
          spreadsheetId:
            existingMember.regiment.spreadsheetId,
          sheetName:
            existingMember.companyName
        });

        await sortCompanyByRank({
          spreadsheetId:
            newRegiment.spreadsheetId,
          sheetName:
            matchedCompany,
          platoon:
            destinationPosition?.type ===
            "platoon"
              ? destinationPosition
              : null
        });

        destinationRow =
          await findMemberRowInCompanyByDiscordId({
            spreadsheetId:
              newRegiment.spreadsheetId,
            sheetName:
              matchedCompany,
            discordId:
              discordMember.id
          }) || destinationRow;

        let updatedNickname = null;
        let nicknameWarning = null;

        try {
          updatedNickname =
            await updateDiscordNickname({
              interaction,
              discordUserId:
                discordMember.id,
              regiment:
                newRegiment,
              rank:
                finalRank,
              robloxUsername:
                memberRecord.robloxUsername
            });
        } catch (nicknameError) {
          console.error(
            "Transfer completed, but nickname updating failed:"
          );
          console.error(nicknameError);

          nicknameWarning =
            nicknameError?.message ||
            "The Discord nickname could not be updated.";
        }

        const replyLines = [
          "Member transferred successfully.",
          "",
          `**Discord Member:** ${discordMember}`,
          `**Discord ID:** ${discordMember.id}`,
          `**Roblox Username:** ${memberRecord.robloxUsername || "Not set"}`,
          `**Previous Rank:** ${previousRank || "Not set"}`,
          `**New Rank:** ${finalRank || "Not set"}`,
          `**Timezone:** ${memberRecord.timezone || "Not set"}`,
          "",
          `**From Regiment:** ${existingMember.regiment.displayName}`,
          `**From Company:** ${existingMember.companyName}`,
          `**From Row:** ${existingMember.row}`,
          "",
          `**To Regiment:** ${newRegiment.displayName}`,
          `**To Company:** ${matchedCompany}`,
          `**To Row:** ${destinationRow}`,
          "",
          "The original ORBAT entry was cleared after the destination entry was created."
        ];

        if (updatedNickname) {
          replyLines.push(
            `**Discord Nickname:** ${updatedNickname}`
          );
        }

        if (nicknameWarning) {
          replyLines.push(
            "",
            "⚠️ The transfer succeeded, but the Discord nickname could not be updated.",
            `**Nickname Error:** ${nicknameWarning}`,
            "Make sure the bot has the Manage Nicknames permission and its role is above the member's highest role."
          );
        }

        if (timezoneWarning) {
          replyLines.push(
            "",
            "⚠️ The transfer succeeded, but timezone processing returned a warning:",
            timezoneWarning
          );
        }

        await sendOrbatLog({
          interaction,
          category: "Member Management",
          action: "Member Transferred",
          affectedMember: discordMember,
          robloxUsername:
            memberRecord.robloxUsername,
          changes: [
            {
              label: "Regiment",
              before:
                existingMember.regiment.displayName,
              after:
                newRegiment.displayName
            },
            {
              label: "Company",
              before:
                existingMember.companyName,
              after:
                matchedCompany
            },
            {
              label: "Rank",
              before: previousRank,
              after: finalRank
            },
            {
              label: "Spreadsheet Row",
              before: existingMember.row,
              after: destinationRow
            },
            {
              label: "Timezone",
              after:
                memberRecord.timezone
            }
          ],
          notes: [
            nicknameWarning
              ? "Discord nickname update failed."
              : null,
            timezoneWarning
              ? "Timezone processing returned a warning."
              : null
          ].filter(Boolean).join(" ") || null
        });

        await interaction.editReply(
          replyLines.join("\n")
        );
      } catch (error) {
        console.error(
          "Failed to transfer member:"
        );
        console.error(error);

        let errorMessage =
          error?.message ||
          "An unknown error occurred.";

        if (errorMessage === "COMPANY_FULL") {
          errorMessage =
            "The destination company has no available member rows.";
        } else if (
          errorMessage === "UNKNOWN_REGIMENT"
        ) {
          errorMessage =
            "The destination regiment could not be recognized.";
        } else if (
          errorMessage === "FIRST_KRUMPER_REKRUT_ONLY"
        ) {
          errorMessage =
            "1. Krümper Kompanie is recruit-only. Only Rekrut may be assigned there.";
        }

        try {
          if (
            interaction.deferred ||
            interaction.replied
          ) {
            await interaction.editReply(
              "Failed to transfer the member: " +
              errorMessage
            );
          } else {
            await interaction.reply({
              content:
                "Failed to transfer the member: " +
                errorMessage,
              flags: MessageFlags.Ephemeral
            });
          }
        } catch (replyError) {
          console.error(
            "Failed to send transfer error reply:"
          );
          console.error(replyError);
        }
      }

      return;
    }

    if (interaction.commandName === "rank") {
      try {
        await interaction.deferReply({
          flags: MessageFlags.Ephemeral
        });

        const discordMember =
          interaction.options.getUser(
            "discord_member",
            true
          );

        const newRank = interaction.options
          .getString("new_rank", true)
          .trim();

        const existingMember =
          await findMemberByDiscordId(
            discordMember.id
          );

        if (!existingMember) {
          await interaction.editReply(
            [
              "That Discord member was not found in the Grand ORBAT.",
              "",
              `**Discord Member:** ${discordMember}`,
              `**Discord ID:** ${discordMember.id}`,
              "",
              "No rank was changed."
            ].join("\n")
          );
          return;
        }

        const memberRecord =
          await getMemberRecord({
            spreadsheetId:
              existingMember.regiment.spreadsheetId,
            sheetName:
              existingMember.companyName,
            row:
              existingMember.row
          });

        const previousRank =
          String(memberRecord.rank || "").trim();

        if (
          normalizeText(previousRank) ===
          normalizeText(newRank)
        ) {
          await interaction.editReply(
            [
              "That member already has the selected rank.",
              "",
              `**Discord Member:** ${discordMember}`,
              `**Regiment:** ${existingMember.regiment.displayName}`,
              `**Company:** ${existingMember.companyName}`,
              `**Rank:** ${newRank}`,
              "",
              "No spreadsheet cells were changed."
            ].join("\n")
          );
          return;
        }

        /*
         * Special recruit promotion:
         * /rank is the ONLY command that automatically moves a Rekrut out
         * of 1. Krümper Kompanie. The bot finds the first Musketier company
         * in the same regiment with a free slot.
         */
        const mustLeaveRecruitCompany =
          isFirstKrumperCompany(
            existingMember.companyName
          ) &&
          isRekrutRank(previousRank) &&
          !isRekrutRank(newRank);

        const mustEnterRecruitCompany =
          !isFirstKrumperCompany(
            existingMember.companyName
          ) &&
          !isRekrutRank(previousRank) &&
          isRekrutRank(newRank);

        if (mustEnterRecruitCompany) {
          const destinationCompany =
            await findFirstKrumperCompany(
              existingMember.regiment
            );

          let destinationRow =
            await addMemberToSheet({
              spreadsheetId:
                existingMember.regiment.spreadsheetId,
              sheetName:
                destinationCompany,
              robloxUsername:
                memberRecord.robloxUsername,
              discordId:
                memberRecord.discordId ||
                discordMember.id,
              rank:
                newRank,
              timezone:
                memberRecord.timezone
            });

          let timezoneWarning = null;

          if (memberRecord.timezone) {
            try {
              await processTimezoneWithAppsScript({
                spreadsheetId:
                  existingMember.regiment.spreadsheetId,
                sheetName:
                  destinationCompany,
                row:
                  destinationRow,
                timezone:
                  memberRecord.timezone
              });
            } catch (webhookError) {
              console.error(
                "Rekrut reassignment succeeded, but timezone processing failed:"
              );
              console.error(webhookError);

              timezoneWarning =
                webhookError?.message ||
                "The Apps Script webhook failed.";
            }
          }

          try {
            await removeMemberFromSheet({
              spreadsheetId:
                existingMember.regiment.spreadsheetId,
              sheetName:
                existingMember.companyName,
              row:
                existingMember.row
            });
          } catch (sourceClearError) {
            try {
              await removeMemberFromSheet({
                spreadsheetId:
                  existingMember.regiment.spreadsheetId,
                sheetName:
                  destinationCompany,
                row:
                  destinationRow
              });
            } catch (rollbackError) {
              console.error(
                "Rekrut reassignment rollback failed:"
              );
              console.error(rollbackError);
            }

            throw sourceClearError;
          }

          await sortCompanyByRank({
            spreadsheetId:
              existingMember.regiment.spreadsheetId,
            sheetName:
              existingMember.companyName
          });

          await sortCompanyByRank({
            spreadsheetId:
              existingMember.regiment.spreadsheetId,
            sheetName:
              destinationCompany
          });

          destinationRow =
            await findMemberRowInCompanyByDiscordId({
              spreadsheetId:
                existingMember.regiment.spreadsheetId,
              sheetName:
                destinationCompany,
              discordId:
                discordMember.id
            }) || destinationRow;

          await writeKrumperEntryDate({
            spreadsheetId:
              existingMember.regiment.spreadsheetId,
            sheetName:
              destinationCompany,
            row:
              destinationRow
          });

          let updatedNickname = null;
          let nicknameWarning = null;

          try {
            updatedNickname =
              await updateDiscordNickname({
                interaction,
                discordUserId:
                  discordMember.id,
                regiment:
                  existingMember.regiment,
                rank:
                  newRank,
                robloxUsername:
                  memberRecord.robloxUsername
              });
          } catch (nicknameError) {
            console.error(
              "Rekrut reassignment succeeded, but nickname updating failed:"
            );
            console.error(nicknameError);

            nicknameWarning =
              nicknameError?.message ||
              "The Discord nickname could not be updated.";
          }

          await sendOrbatLog({
            interaction,
            category: "Rank Management",
            action:
              "Member Changed to Rekrut and Automatically Transferred",
            affectedMember:
              discordMember,
            robloxUsername:
              memberRecord.robloxUsername,
            changes: [
              {
                label: "Regiment",
                after:
                  existingMember.regiment.displayName
              },
              {
                label: "Company",
                before:
                  existingMember.companyName,
                after:
                  destinationCompany
              },
              {
                label: "Rank",
                before:
                  previousRank,
                after:
                  newRank
              },
              {
                label: "Spreadsheet Row",
                before:
                  existingMember.row,
                after:
                  destinationRow
              }
            ],
            notes: [
              "Automatic transfer to 1. Krümper-Kompanie triggered by /rank.",
              nicknameWarning
                ? "Discord nickname update failed."
                : null,
              timezoneWarning
                ? "Timezone processing returned a warning."
                : null
            ].filter(Boolean).join(" ")
          });

          const replyLines = [
            "Member rank updated successfully.",
            "",
            "Because the new rank is Rekrut, the member was automatically transferred to 1. Krümper-Kompanie.",
            "",
            `**Discord Member:** ${discordMember}`,
            `**Roblox Username:** ${memberRecord.robloxUsername || "Not set"}`,
            `**Regiment:** ${existingMember.regiment.displayName}`,
            `**Previous Company:** ${existingMember.companyName}`,
            `**New Company:** ${destinationCompany}`,
            `**Previous Rank:** ${previousRank}`,
            `**New Rank:** ${newRank}`,
            `**New Spreadsheet Row:** ${destinationRow}`
          ];

          if (updatedNickname) {
            replyLines.push(
              `**Discord Nickname:** ${updatedNickname}`
            );
          }

          if (nicknameWarning) {
            replyLines.push(
              "",
              "⚠️ The rank change succeeded, but the Discord nickname could not be updated.",
              `**Nickname Error:** ${nicknameWarning}`
            );
          }

          if (timezoneWarning) {
            replyLines.push(
              "",
              "⚠️ The rank change succeeded, but timezone processing returned a warning:",
              timezoneWarning
            );
          }

          await interaction.editReply(
            replyLines.join("\n")
          );

          return;
        }

        if (mustLeaveRecruitCompany) {
          const destinationCompany =
            await findAvailableMusketierCompany(
              existingMember.regiment
            );

          let destinationRow =
            await addMemberToSheet({
              spreadsheetId:
                existingMember.regiment.spreadsheetId,
              sheetName:
                destinationCompany,
              robloxUsername:
                memberRecord.robloxUsername,
              discordId:
                memberRecord.discordId ||
                discordMember.id,
              rank:
                newRank,
              timezone:
                memberRecord.timezone
            });

          let timezoneWarning = null;

          if (memberRecord.timezone) {
            try {
              await processTimezoneWithAppsScript({
                spreadsheetId:
                  existingMember.regiment.spreadsheetId,
                sheetName:
                  destinationCompany,
                row:
                  destinationRow,
                timezone:
                  memberRecord.timezone
              });
            } catch (webhookError) {
              console.error(
                "Recruit promotion transfer succeeded, but timezone processing failed:"
              );
              console.error(webhookError);

              timezoneWarning =
                webhookError?.message ||
                "The Apps Script webhook failed.";
            }
          }

          try {
            await removeMemberFromSheet({
              spreadsheetId:
                existingMember.regiment.spreadsheetId,
              sheetName:
                existingMember.companyName,
              row:
                existingMember.row
            });
          } catch (sourceClearError) {
            try {
              await removeMemberFromSheet({
                spreadsheetId:
                  existingMember.regiment.spreadsheetId,
                sheetName:
                  destinationCompany,
                row:
                  destinationRow
              });
            } catch (rollbackError) {
              console.error(
                "Recruit promotion rollback failed:"
              );
              console.error(rollbackError);
            }

            throw sourceClearError;
          }

          await sortCompanyByRank({
            spreadsheetId:
              existingMember.regiment.spreadsheetId,
            sheetName:
              existingMember.companyName
          });

          await sortCompanyByRank({
            spreadsheetId:
              existingMember.regiment.spreadsheetId,
            sheetName:
              destinationCompany
          });

          destinationRow =
            await findMemberRowInCompanyByDiscordId({
              spreadsheetId:
                existingMember.regiment.spreadsheetId,
              sheetName:
                destinationCompany,
              discordId:
                discordMember.id
            }) || destinationRow;

          let updatedNickname = null;
          let nicknameWarning = null;

          try {
            updatedNickname =
              await updateDiscordNickname({
                interaction,
                discordUserId:
                  discordMember.id,
                regiment:
                  existingMember.regiment,
                rank:
                  newRank,
                robloxUsername:
                  memberRecord.robloxUsername
              });
          } catch (nicknameError) {
            console.error(
              "Recruit promotion succeeded, but nickname updating failed:"
            );
            console.error(nicknameError);

            nicknameWarning =
              nicknameError?.message ||
              "The Discord nickname could not be updated.";
          }

          await sendOrbatLog({
            interaction,
            category: "Rank Management",
            action:
              "Rekrut Promoted and Automatically Transferred",
            affectedMember: discordMember,
            robloxUsername:
              memberRecord.robloxUsername,
            changes: [
              {
                label: "Regiment",
                after:
                  existingMember.regiment.displayName
              },
              {
                label: "Company",
                before:
                  existingMember.companyName,
                after:
                  destinationCompany
              },
              {
                label: "Rank",
                before:
                  previousRank,
                after:
                  newRank
              },
              {
                label: "Spreadsheet Row",
                before:
                  existingMember.row,
                after:
                  destinationRow
              }
            ],
            notes: [
              "Automatic Musketier transfer triggered by /rank.",
              nicknameWarning
                ? "Discord nickname update failed."
                : null,
              timezoneWarning
                ? "Timezone processing returned a warning."
                : null
            ].filter(Boolean).join(" ")
          });

          const replyLines = [
            "Member promoted successfully.",
            "",
            "Because the member was a Rekrut in 1. Krümper Kompanie, they were automatically transferred to an available Musketier company.",
            "",
            `**Discord Member:** ${discordMember}`,
            `**Roblox Username:** ${memberRecord.robloxUsername || "Not set"}`,
            `**Regiment:** ${existingMember.regiment.displayName}`,
            `**Previous Company:** ${existingMember.companyName}`,
            `**New Company:** ${destinationCompany}`,
            `**Previous Rank:** ${previousRank}`,
            `**New Rank:** ${newRank}`,
            `**New Spreadsheet Row:** ${destinationRow}`
          ];

          if (updatedNickname) {
            replyLines.push(
              `**Discord Nickname:** ${updatedNickname}`
            );
          }

          if (nicknameWarning) {
            replyLines.push(
              "",
              "⚠️ The promotion succeeded, but the Discord nickname could not be updated.",
              `**Nickname Error:** ${nicknameWarning}`
            );
          }

          if (timezoneWarning) {
            replyLines.push(
              "",
              "⚠️ The promotion succeeded, but timezone processing returned a warning:",
              timezoneWarning
            );
          }

          await interaction.editReply(
            replyLines.join("\n")
          );

          return;
        }

        /*
         * Normal rank change.
         */
        assertCompanyRankAllowed(
          existingMember.companyName,
          newRank
        );

        await updateMemberRank({
          spreadsheetId:
            existingMember.regiment.spreadsheetId,
          sheetName:
            existingMember.companyName,
          row:
            existingMember.row,
          rank:
            newRank
        });

        await sortCompanyByRank({
          spreadsheetId:
            existingMember.regiment.spreadsheetId,
          sheetName:
            existingMember.companyName
        });

        const sortedRow =
          await findMemberRowInCompanyByDiscordId({
            spreadsheetId:
              existingMember.regiment.spreadsheetId,
            sheetName:
              existingMember.companyName,
            discordId:
              discordMember.id
          }) || existingMember.row;

        let updatedNickname = null;
        let nicknameWarning = null;

        try {
          updatedNickname =
            await updateDiscordNickname({
              interaction,
              discordUserId:
                discordMember.id,
              regiment:
                existingMember.regiment,
              rank:
                newRank,
              robloxUsername:
                memberRecord.robloxUsername
            });
        } catch (nicknameError) {
          console.error(
            "Rank changed, but nickname updating failed:"
          );
          console.error(nicknameError);

          nicknameWarning =
            nicknameError?.message ||
            "The Discord nickname could not be updated.";
        }

        const rankReplyLines = [
          "Member rank updated successfully.",
          "",
          `**Discord Member:** ${discordMember}`,
          `**Discord ID:** ${discordMember.id}`,
          `**Regiment:** ${existingMember.regiment.displayName}`,
          `**Company:** ${existingMember.companyName}`,
          `**Previous Rank:** ${previousRank}`,
          `**New Rank:** ${newRank}`,
          `**Spreadsheet Row After Sorting:** ${sortedRow}`
        ];

        if (updatedNickname) {
          rankReplyLines.push(
            `**Discord Nickname:** ${updatedNickname}`
          );
        }

        if (nicknameWarning) {
          rankReplyLines.push(
            "",
            "⚠️ The rank changed, but the Discord nickname could not be updated.",
            `**Nickname Error:** ${nicknameWarning}`,
            "Make sure the bot has the Manage Nicknames permission and its role is above the member's highest role."
          );
        }

        await sendOrbatLog({
          interaction,
          category: "Rank Management",
          action: "Rank Changed",
          affectedMember: discordMember,
          robloxUsername:
            memberRecord.robloxUsername,
          changes: [
            {
              label: "Regiment",
              after:
                existingMember.regiment.displayName
            },
            {
              label: "Company",
              after:
                existingMember.companyName
            },
            {
              label: "Rank",
              before:
                previousRank,
              after:
                newRank
            },
            {
              label: "Spreadsheet Row",
              before:
                existingMember.row,
              after:
                sortedRow
            }
          ],
          notes: nicknameWarning
            ? "Rank updated, but the Discord nickname could not be updated."
            : "Company automatically re-sorted by rank."
        });

        await interaction.editReply(
          rankReplyLines.join("\n")
        );

      } catch (error) {
        console.error(
          "Failed to change member rank:"
        );
        console.error(error);

        let errorMessage =
          error?.message ||
          "An unknown error occurred.";

        if (
          errorMessage ===
          "FIRST_KRUMPER_REKRUT_ONLY"
        ) {
          errorMessage =
            "1. Krümper Kompanie is recruit-only. Members in that company must have the rank Rekrut.";
        } else if (
          errorMessage ===
          "NO_MUSKETIER_COMPANY"
        ) {
          errorMessage =
            "The regiment does not have a Musketier company available for the automatic recruit transfer.";
        } else if (
          errorMessage ===
          "MUSKETIER_COMPANIES_FULL"
        ) {
          errorMessage =
            "The member was not promoted because every Musketier company in the regiment is full.";
        } else if (
          errorMessage ===
          "FIRST_KRUMPER_COMPANY_NOT_FOUND"
        ) {
          errorMessage =
            "The regiment does not have a 1. Krümper-Kompanie sheet configured.";
        } else if (
          errorMessage ===
          "FIRST_KRUMPER_COMPANY_FULL"
        ) {
          errorMessage =
            "The rank change was not made because 1. Krümper-Kompanie has no open member slots.";
        } else if (
          errorMessage ===
          "REKRUT_FIRST_KRUMPER_ONLY"
        ) {
          errorMessage =
            "Members with the rank Rekrut may only be assigned to 1. Krümper-Kompanie.";
        }

        try {
          if (
            interaction.deferred ||
            interaction.replied
          ) {
            await interaction.editReply(
              "Failed to update the member's rank: " +
              errorMessage
            );
          } else {
            await interaction.reply({
              content:
                "Failed to update the member's rank: " +
                errorMessage,
              flags: MessageFlags.Ephemeral
            });
          }
        } catch (replyError) {
          console.error(
            "Failed to send promotion error reply:"
          );
          console.error(replyError);
        }
      }

      return;
    }

    if (interaction.commandName === "removemember") {
      try {
        await interaction.deferReply({
          flags: MessageFlags.Ephemeral
        });

        const discordMember =
          interaction.options.getUser(
            "discord_member",
            true
          );

        const existingMember =
          await findMemberByDiscordId(
            discordMember.id
          );

        if (!existingMember) {
          await interaction.editReply(
            [
              "That Discord member was not found in the Grand ORBAT.",
              "",
              `**Discord Member:** ${discordMember}`,
              `**Discord ID:** ${discordMember.id}`,
              "",
              "No spreadsheet cells were changed."
            ].join("\n")
          );
          return;
        }

        /*
         * findMemberByDiscordId only returns the member's location.
         * Read C:F before clearing the row so the Roblox username is preserved.
         */
        const memberRecord =
          await getMemberRecord({
            spreadsheetId:
              existingMember.regiment.spreadsheetId,
            sheetName:
              existingMember.companyName,
            row:
              existingMember.row
          });

        await removeMemberFromSheet({
          spreadsheetId:
            existingMember.regiment.spreadsheetId,
          sheetName:
            existingMember.companyName,
          row:
            existingMember.row
        });

        await sortCompanyByRank({
          spreadsheetId:
            existingMember.regiment.spreadsheetId,
          sheetName:
            existingMember.companyName
        });

        let nicknameResult;

        try {
          const updatedNickname =
            await resetDiscordNicknameToRobloxUsername({
              interaction,
              discordUserId:
                discordMember.id,
              robloxUsername:
                memberRecord.robloxUsername
            });

          nicknameResult =
            `**Nickname:** ${updatedNickname}`;
        } catch (nicknameError) {
          console.error(
            "Member was removed, but nickname resetting failed:"
          );
          console.error(nicknameError);

          nicknameResult =
            "**Nickname:** Could not be changed. Check that the Roblox username was stored in column C and that the bot has Manage Nicknames permission.";
        }

        await sendOrbatLog({
          interaction,
          category: "Member Management",
          action: "Member Removed",
          affectedMember: discordMember,
          robloxUsername:
            memberRecord.robloxUsername,
          changes: [
            {
              label: "Regiment",
              before:
                existingMember.regiment.displayName,
              after: "Removed from ORBAT"
            },
            {
              label: "Company",
              before:
                existingMember.companyName,
              after: "Removed from ORBAT"
            },
            {
              label: "Rank",
              before:
                memberRecord.rank,
              after: "Removed from ORBAT"
            },
            {
              label: "Timezone",
              before:
                memberRecord.timezone,
              after: "Removed from ORBAT"
            },
            {
              label: "Spreadsheet Row",
              before: existingMember.row,
              after: "Cleared"
            }
          ],
          notes:
            nicknameResult.includes("Could not")
              ? "Member removed, but the Discord nickname could not be reset."
              : "Discord nickname was reset to the Roblox username."
        });

        await interaction.editReply(
          [
            "Member removed successfully.",
            "",
            `**Discord Member:** ${discordMember}`,
            `**Discord ID:** ${discordMember.id}`,
            `**Roblox Username:** ${memberRecord.robloxUsername || "Not set"}`,
            `**Regiment:** ${existingMember.regiment.displayName}`,
            `**Company:** ${existingMember.companyName}`,
            `**Spreadsheet Row:** ${existingMember.row}`,
            nicknameResult,
            "",
            `Cleared C${existingMember.row}, D${existingMember.row}, E${existingMember.row}, and F${existingMember.row}.`
          ].join("\n")
        );
      } catch (error) {
        console.error("Failed to remove member:");
        console.error(error);

        const errorMessage =
          error?.message ||
          "An unknown error occurred.";

        try {
          if (
            interaction.deferred ||
            interaction.replied
          ) {
            await interaction.editReply(
              "Failed to remove the member: " +
              errorMessage
            );
          } else {
            await interaction.reply({
              content:
                "Failed to remove the member: " +
                errorMessage,
              flags: MessageFlags.Ephemeral
            });
          }
        } catch (replyError) {
          console.error(
            "Failed to send removal error reply:"
          );
          console.error(replyError);
        }
      }

      return;
    }

    try {
      await interaction.deferReply({
        flags: MessageFlags.Ephemeral
      });
      const discordMember =
        interaction.options.getUser(
          "discord_member",
          true
        );

      const guildId = interaction.guildId;

      if (!guildId) {
        await interaction.editReply(
          "This command can only be used inside a Discord server."
        );
        return;
      }

      let verifiedRobloxAccount;

      try {
        verifiedRobloxAccount =
          await getVerifiedRobloxAccount({
            guildId,
            discordId:
              discordMember.id
          });
      } catch (roverError) {
        const roverCode =
          roverError?.message || "";

        if (
          roverCode ===
          "ROVER_NOT_VERIFIED"
        ) {
          await interaction.editReply(
            [
              "RoVer could not find a verified Roblox account for that member.",
              "",
              `**Discord Member:** ${discordMember}`,
              `**Discord ID:** ${discordMember.id}`,
              "",
              "Ask the member to verify with RoVer and then run this command again."
            ].join("\n")
          );
          return;
        }

        if (
          roverCode ===
          "ROVER_ACCESS_DENIED"
        ) {
          await interaction.editReply(
            [
              "RoVer would not allow this bot to access that member's Roblox account.",
              "",
              "Make sure:",
              "1. `ROVER_API_KEY` is correct in the `.env` file.",
              "2. The API key belongs to this Discord server.",
              "3. The member has granted the server and other bots permission through RoVer's `/privacy` command."
            ].join("\n")
          );
          return;
        }

        if (
          roverCode ===
          "ROVER_INVALID_RESPONSE"
        ) {
          await interaction.editReply(
            "RoVer returned an account response that this bot could not understand."
          );
          return;
        }

        throw roverError;
      }

      const robloxUsername =
        verifiedRobloxAccount.robloxUsername;

      const regimentValue = interaction.options
        .getString("regiment", true)
        .trim();

      const company = interaction.options
        .getString("company", true)
        .trim();

      const rank = interaction.options
        .getString("rank", true)
        .trim();

      const timezone = interaction.options
        .getString("timezone", true)
        .trim();

      const positionValue =
        interaction.options
          .getString(
            "position",
            false
          )
          ?.trim() || null;

      if (!rank) {
        await interaction.editReply(
          "The rank cannot be empty."
        );
        return;
      }

      if (!timezone) {
        await interaction.editReply(
          "The timezone cannot be empty."
        );
        return;
      }

      const regiment = resolveRegiment(regimentValue);

      let schuetzenPosition =
        null;

      if (
        isSchuetzenRegiment(
          regiment
        )
      ) {
        if (!positionValue) {
          await interaction.editReply(
            [
              "The member was not added.",
              "",
              "**Schlesisches Schützen-Bataillon uses a position system.**",
              "Select **Company Commander**, **1. Platoon**, or **2. Platoon** in the `position` option and try again."
            ].join("\n")
          );
          return;
        }

        try {
          schuetzenPosition =
            resolveSchuetzenPosition(
              positionValue
            );
        } catch {
          await interaction.editReply(
            "That Schützen position is not configured."
          );
          return;
        }
      }

      const availableCompanies =
        await getCompanySheetNames(
          regiment
        );

      const matchedCompany =
        availableCompanies.find(
          availableCompany =>
            normalizeText(availableCompany) ===
            normalizeText(company)
        );

      if (!matchedCompany) {
        await interaction.editReply(
          [
            "The selected company could not be used.",
            "",
            `**Regiment:** ${regiment.displayName}`,
            `**Company Submitted:** ${company}`,
            "",
            "The regiment overview sheet and administrative sheets cannot contain members.",
            "Select an actual company from autocomplete and try again."
          ].join("\n")
        );
        return;
      }

      console.log("REKRUT VALIDATION:", {
        rank,
        company: matchedCompany,
        normalizedRank:
          normalizeText(rank),
        normalizedCompany:
          normalizeText(matchedCompany),
        isRekrut:
          isRekrutRank(rank),
        isFirstKrumper:
          isFirstKrumperCompany(
            matchedCompany
          )
      });

      try {
        if (
          !isSchuetzenRegiment(
            regiment
          )
        ) {
          assertCompanyRankAllowed(
            matchedCompany,
            rank
          );
        }
      } catch (companyRankError) {
        if (
          companyRankError?.message ===
          "REKRUT_FIRST_KRUMPER_ONLY"
        ) {
          await interaction.editReply(
            [
              "The member was not added.",
              "",
              `**Selected Rank:** ${rank}`,
              `**Selected Company:** ${matchedCompany}`,
              "",
              "That company is not **1. Krümper-Kompanie**.",
              "Members with the rank **Rekrut** can only be added to **1. Krümper-Kompanie**.",
              "",
              "Select **1. Krümper-Kompanie** and try again."
            ].join("\n")
          );
          return;
        }

        if (
          companyRankError?.message ===
          "FIRST_KRUMPER_REKRUT_ONLY"
        ) {
          await interaction.editReply(
            [
              "The member was not added.",
              "",
              `**Selected Rank:** ${rank}`,
              `**Selected Company:** ${matchedCompany}`,
              "",
              "**1. Krümper-Kompanie** can only contain members with the rank **Rekrut**."
            ].join("\n")
          );
          return;
        }

        throw companyRankError;
      }

      const existingMember =
        await findMemberByDiscordId(
          discordMember.id
        );

      if (existingMember) {
        await interaction.editReply(
          [
            "This Discord member is already listed in the Grand ORBAT.",
            "",
            `**Discord Member:** ${discordMember}`,
            `**Discord ID:** ${discordMember.id}`,
            `**Existing Regiment:** ${existingMember.regiment.displayName}`,
            `**Existing Company:** ${existingMember.companyName}`,
            `**Existing Row:** ${existingMember.row}`,
            "",
            "The member was not added again."
          ].join("\n")
        );
        return;
      }

      let row = await addMemberToSheet({
        spreadsheetId: regiment.spreadsheetId,
        sheetName: matchedCompany,
        robloxUsername,
        discordId: discordMember.id,
        rank,
        timezone,
        position:
          schuetzenPosition
      });

      let timezoneResult = null;
      let timezoneWarning = null;

      try {
        timezoneResult =
          await processTimezoneWithAppsScript({
            spreadsheetId: regiment.spreadsheetId,
            sheetName: matchedCompany,
            row,
            timezone
          });
      } catch (webhookError) {
        console.error(
          "Member was added, but timezone processing failed:"
        );
        console.error(webhookError);

        timezoneWarning =
          webhookError?.message ||
          "The Apps Script webhook failed.";
      }

      await sortCompanyByRank({
        spreadsheetId:
          regiment.spreadsheetId,
        sheetName:
          matchedCompany,
        platoon:
          schuetzenPosition?.type ===
          "platoon"
            ? schuetzenPosition
            : null
      });

      row =
        await findMemberRowInCompanyByDiscordId({
          spreadsheetId:
            regiment.spreadsheetId,
          sheetName:
            matchedCompany,
          discordId:
            discordMember.id
        }) || row;

      await writeKrumperEntryDate({
        spreadsheetId:
          regiment.spreadsheetId,
        sheetName:
          matchedCompany,
        row
      });

      let nicknameResult = null;
      let nicknameWarning = null;

      try {
        nicknameResult =
          await updateDiscordNickname({
            interaction,
            discordUserId:
              discordMember.id,
            regiment,
            rank,
            robloxUsername
          });
      } catch (nicknameError) {
        console.error(
          "Member was added, but nickname updating failed:"
        );
        console.error(nicknameError);

        nicknameWarning =
          nicknameError?.message ||
          "The Discord nickname could not be updated.";
      }

      const replyLines = [
        "Member added successfully.",
        "",
        `**Regiment:** ${regiment.displayName}`,
        `**Roblox Username:** ${robloxUsername}`,
        `**Roblox ID:** ${verifiedRobloxAccount.robloxId || "Not provided by RoVer"}`,
        `**Discord Member:** ${discordMember}`,
        `**Discord ID:** ${discordMember.id}`,
        `**Company:** ${matchedCompany}`,
        `**Rank:** ${rank}`,
        `**Timezone Submitted:** ${timezone}`,
        `**Spreadsheet Row:** ${row}`,
        "",
        `Written to C${row}, D${row}, E${row}, and F${row}.`
      ];

      if (nicknameResult) {
        replyLines.push(
          `**Discord Nickname:** ${nicknameResult}`
        );
      }

      if (nicknameWarning) {
        replyLines.push(
          "",
          "⚠️ The member was added, but the Discord nickname could not be updated.",
          `**Nickname Error:** ${nicknameWarning}`,
          "Make sure the bot has the Manage Nicknames permission and its role is above the member's highest role."
        );
      }

      if (timezoneResult?.displayValue) {
        replyLines.push(
          `**Processed Timezone:** ` +
          `${timezoneResult.displayValue}`
        );
      }

      if (timezoneResult?.storageValue) {
        replyLines.push(
          `**Stored IANA Timezone:** ` +
          `${timezoneResult.storageValue}`
        );
      }

      if (timezoneWarning) {
        replyLines.push(
          "",
          "⚠️ The member was added, but the timezone could not be processed automatically.",
          `**Webhook Error:** ${timezoneWarning}`,
          `The original timezone remains in F${row}.`
        );
      }

      await sendOrbatLog({
        interaction,
        category: "Member Management",
        action: "Member Added",
        affectedMember: discordMember,
        robloxUsername,
        changes: [
          {
            label: "Regiment",
            after: regiment.displayName
          },
          {
            label: "Company",
            after: matchedCompany
          },
          {
            label: "Rank",
            after: rank
          },
          {
            label: "Timezone",
            after: timezone
          },
          {
            label: "Spreadsheet Row",
            after: row
          },
          {
            label: "Discord ID",
            after: discordMember.id
          }
        ],
        notes: [
          nicknameWarning
            ? "Discord nickname update failed."
            : null,
          timezoneWarning
            ? "Timezone processing returned a warning."
            : null
        ].filter(Boolean).join(" ") || null
      });

      await interaction.editReply(
        replyLines.join("\n")
      );
    } catch (error) {
      console.error("Failed to add member:");
      console.error(error);

      if (
        error?.code === 10062 ||
        error?.code === 40060
      ) {
        console.error(
          "The Discord interaction expired or was already acknowledged. " +
          "Reset the bot token if another remote instance may still be running."
        );
        return;
      }

      let errorMessage =
        "The member could not be added to the spreadsheet.";

      if (error?.message === "UNKNOWN_REGIMENT") {
        errorMessage =
          "The selected regiment is not configured. Re-register the slash command and make sure its regiment values are 11_schlesisches, 6_westpreussisches, or ostpreussisches_jaeger.";
      } else if (error?.message === "COMPANY_FULL") {
        const selectedCompany =
          interaction.options
            .getString("company", true)
            .trim();

        const lastMemberRow =
          getLastMemberRow(selectedCompany);

        errorMessage =
          `The selected company roster is full. ` +
          `There are no empty member slots between rows ` +
          `${FIRST_MEMBER_ROW} and ${lastMemberRow}.`;
      } else if (
        error?.code === 404 ||
        error?.response?.status === 404
      ) {
        errorMessage =
          "The regiment spreadsheet or selected company worksheet could not be found. Verify the spreadsheet IDs and ensure the company name exactly matches the Google Sheets tab.";
      } else if (
        error?.code === 403 ||
        error?.response?.status === 403
      ) {
        errorMessage =
          "The bot does not have permission to edit the selected regiment spreadsheet. Share all three spreadsheets with the client_email listed in service-account.json and give it Editor access.";
      } else if (
        String(error?.message || "").includes(
          "Unable to parse range"
        )
      ) {
        errorMessage =
          "The selected company name does not exactly match a worksheet tab in the selected regiment spreadsheet.";
      } else if (
        String(error?.message || "").includes(
          "Requested entity was not found"
        )
      ) {
        errorMessage =
          "Google could not find the selected regiment spreadsheet. Check the three spreadsheet IDs in your .env file.";
      }

      try {
        if (interaction.deferred || interaction.replied) {
          await interaction.editReply(errorMessage);
        } else {
          await interaction.reply({
            content: errorMessage,
            flags: MessageFlags.Ephemeral
          });
        }
      } catch (responseError) {
        if (
          responseError?.code !== 10062 &&
          responseError?.code !== 40060
        ) {
          console.error(
            "Could not send the command error response:"
          );
          console.error(responseError);
        }
      }
    }
  }
);

/*
|--------------------------------------------------------------------------
| Error handling
|--------------------------------------------------------------------------
*/

client.on(Events.Error, error => {
  console.error("Discord client error:");
  console.error(error);
});

process.on("unhandledRejection", error => {
  console.error("Unhandled promise rejection:");
  console.error(error);
});

process.on("uncaughtException", error => {
  console.error("Uncaught exception:");
  console.error(error);
});

/*
|--------------------------------------------------------------------------
| Log in
|--------------------------------------------------------------------------
*/

client.login(DISCORD_TOKEN);