import "dotenv/config";

import fs from "node:fs";
import {
  Client,
  EmbedBuilder,
  Events,
  GatewayIntentBits,
  MessageFlags
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
  ["JAEGER_SPREADSHEET_ID", JAEGER_SPREADSHEET_ID]
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

  const range =
    `${safeSheetName}!H${row}`;

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[dateValue]]
    }
  });

  let verifyResponse =
    await sheets.spreadsheets.values.get({
      spreadsheetId,
      range
    });

  let storedDate = String(
    verifyResponse.data.values?.[0]?.[0] || ""
  ).trim();

  if (!storedDate) {
    console.warn(
      "KRUMPER ENTRY DATE WAS BLANK AFTER FIRST WRITE. RETRYING:",
      {
        sheetName,
        row,
        date: dateValue
      }
    );

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range,
      valueInputOption: "RAW",
      requestBody: {
        values: [[dateValue]]
      }
    });

    verifyResponse =
      await sheets.spreadsheets.values.get({
        spreadsheetId,
        range
      });

    storedDate = String(
      verifyResponse.data.values?.[0]?.[0] || ""
    ).trim();
  }

  if (!storedDate) {
    throw new Error(
      `KRUMPER_DATE_WRITE_FAILED: Google Sheets returned a blank value for ${sheetName}!H${row}.`
    );
  }

  console.log(
    "KRUMPER ENTRY DATE VERIFIED:",
    {
      sheetName,
      row,
      requestedDate: dateValue,
      storedDate
    }
  );
}

async function verifyKrumperDateCleared({
  spreadsheetId,
  sheetName,
  row
}) {
  const safeSheetName =
    escapeSheetName(sheetName);

  const range =
    `${safeSheetName}!H${row}`;

  const response =
    await sheets.spreadsheets.values.get({
      spreadsheetId,
      range
    });

  const value = String(
    response.data.values?.[0]?.[0] || ""
  ).trim();

  if (value) {
    await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range
    });
  }
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
  const lastMemberRow = getLastMemberRow(sheetName);

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

async function sortCompanyByRank({
  spreadsheetId,
  sheetName
}) {
  const safeSheetName = escapeSheetName(sheetName);
  const lastMemberRow = getLastMemberRow(sheetName);
  const slotCount =
    lastMemberRow - FIRST_MEMBER_ROW + 1;

  /*
   * C:H are kept together so the hidden timezone storage in G and the
   * Krümper entry date in H stay attached to the correct member while rows are reordered.
   */
  const response =
    await sheets.spreadsheets.values.get({
      spreadsheetId,
      range:
        `${safeSheetName}!C${FIRST_MEMBER_ROW}:` +
        `H${lastMemberRow}`,
      majorDimension: "ROWS"
    });

  const members = (response.data.values || [])
    .map(row => {
      const padded = Array.from(
        { length: 6 },
        (_, index) => row[index] ?? ""
      );

      return padded;
    })
    .filter(row =>
      String(row[0] || "").trim() ||
      String(row[1] || "").trim()
    );

  members.sort((a, b) => {
    const rankA =
      RANK_SORT_PRIORITY.get(
        normalizeText(a[2])
      ) ?? 999;

    const rankB =
      RANK_SORT_PRIORITY.get(
        normalizeText(b[2])
      ) ?? 999;

    if (rankA !== rankB) {
      return rankA - rankB;
    }

    return String(a[0] || "").localeCompare(
      String(b[0] || ""),
      undefined,
      { sensitivity: "base" }
    );
  });

  const rewrittenRows =
    Array.from(
      { length: slotCount },
      (_, index) =>
        members[index] || ["", "", "", "", "", ""]
    );

  /*
   * Column H belongs exclusively to 1. Krümper-Kompanie.
   * If this is any other company, strip H from every roster row while
   * preserving C:G. This also cleans up stale dates left by older builds.
   */
  if (!isFirstKrumperCompany(sheetName)) {
    for (const row of rewrittenRows) {
      row[5] = "";
    }
  }

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range:
      `${safeSheetName}!C${FIRST_MEMBER_ROW}:` +
      `H${lastMemberRow}`,
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
        getLastMemberRow(companyName);

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
  const safeSheetName = escapeSheetName(sheetName);

  /*
   * Clear only the fields managed by this bot:
   * C = Roblox username
   * D = Discord ID
   * E = rank
   * F = timezone
   *
   * The row itself is not deleted, preserving sheet formatting and formulas.
   */
  await sheets.spreadsheets.values.batchClear({
    spreadsheetId,
    requestBody: {
      ranges: [
        `${safeSheetName}!C${row}`,
        `${safeSheetName}!D${row}`,
        `${safeSheetName}!E${row}`,
        `${safeSheetName}!F${row}`,
        `${safeSheetName}!H${row}`
      ]
    }
  });

  await verifyKrumperDateCleared({
    spreadsheetId,
    sheetName,
    row
  });

  console.log(
    "ORBAT ROW CLEARED:",
    {
      sheetName,
      row,
      clearedColumns: ["C", "D", "E", "F", "H"]
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
  timezone
}) {
  /*
   * Final hard-lock validation.
   * A Rekrut may only be written to 1. Krümper-Kompanie, and
   * 1. Krümper-Kompanie may only contain Rekruten.
   */
  assertCompanyRankAllowed(
    sheetName,
    rank
  );

  const row = await findFirstEmptyRow({
    spreadsheetId,
    sheetName
  });

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
  if (isFirstKrumperCompany(sheetName)) {
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
  } else {
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

  const companies =
    await getCompanySheetNames(
      regiment
    );

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
      if (
        interaction.commandName !== "addmember" &&
        interaction.commandName !== "transfer"
      ) {
        return;
      }

      try {
        await handleCompanyAutocomplete(
          interaction
        );
      } catch (error) {
        console.error(
          "Company autocomplete failed:"
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

    if (!interaction.isChatInputCommand()) {
      return;
    }

    if (
      interaction.commandName !== "addmember" &&
      interaction.commandName !== "removemember" &&
      interaction.commandName !== "rank" &&
      interaction.commandName !== "transfer"
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

        const sameRegiment =
          existingMember.regiment.spreadsheetId ===
          newRegiment.spreadsheetId;

        const sameCompany =
          normalizeText(
            existingMember.companyName
          ) === normalizeText(newCompany);

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
          assertCompanyRankAllowed(
            matchedCompany,
            finalRank
          );
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
              existingMember.companyName
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
              memberRecord.timezone
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
            matchedCompany
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
        assertCompanyRankAllowed(
          matchedCompany,
          rank
        );
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
        timezone
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
          matchedCompany
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