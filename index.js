
  }

  return (
    MASTER_ROSTER_FIRST_MEMBER_ROW +
    values.length
  );
}


function getMasterRosterPosition({
  regiment,
  company,
  row = null,
  explicitPosition = null
}) {
  if (explicitPosition) {
    if (
      typeof explicitPosition ===
      "string"
    ) {
      try {
        return (
          resolveSchuetzenPosition(
            explicitPosition
          ).displayName
        );
      } catch {
        return String(
          explicitPosition
        ).trim();
      }
    }

    return String(
      explicitPosition.displayName ||
      explicitPosition.key ||
      ""
    ).trim();
  }

  if (isGarnisonCompany(company)) {
    return "Garnison";
  }

  if (
    row &&
    usesSchuetzenPositionStructure(
      regiment,
      company
    )
  ) {
    return (
      getSchuetzenPositionForRow(row)
        ?.displayName || ""
    );
  }

  return "";
}


async function syncMasterRosterMember({
  discordId,
  robloxUsername,
  rank,
  regiment,
  kompanie,
  position = "",
  active = "Yes"
}) {
  const sheetName =
    await getMasterRosterSheetName();

  const safeSheetName =
    escapeSheetName(sheetName);

  const existing =
    await findMasterRosterMember(
      discordId
    );

  const today =
    getCurrentOrbatDate();

  if (!existing) {
    const row =
      await findFirstEmptyMasterRosterRow();

    await sheets.spreadsheets.values.update({
      spreadsheetId:
        MASTER_ROSTER_SPREADSHEET_ID,
      range:
        `${safeSheetName}!B${row}:J${row}`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[
          robloxUsername,
          String(discordId),
          rank,
          today,
          today,
          regiment,
          kompanie,
          position,
          active
        ]]
      }
    });

    console.log(
      "MASTER ROSTER MEMBER CREATED:",
      {
        discordId,
        row,
        rank,
        regiment,
        kompanie,
        position
      }
    );

    return {
      row,
      created: true,
      rankChanged: false
    };
  }

  const rankChanged =
    normalizeText(existing.rank) !==
    normalizeText(rank);

  /*
   * E (Date Joined Army) is deliberately NOT written here.
   * Once a member exists, their original Army join date is permanent.
   *
   * F (Date of Current Rank) changes only when D (Current Rank) changes.
   */
  const updates = [
    {
      range:
        `${safeSheetName}!B${existing.row}`,
      values: [[robloxUsername]]
    },
    {
      range:
        `${safeSheetName}!C${existing.row}`,
      values: [[String(discordId)]]
    },
    {
      range:
        `${safeSheetName}!D${existing.row}`,
      values: [[rank]]
    },
    {
      range:
        `${safeSheetName}!G${existing.row}`,
      values: [[regiment]]
    },
    {
      range:
        `${safeSheetName}!H${existing.row}`,
      values: [[kompanie]]
    },
    {
      range:
        `${safeSheetName}!I${existing.row}`,
      values: [[position]]
    },
    {
      range:
        `${safeSheetName}!J${existing.row}`,
      values: [[active]]
    }
  ];

  if (rankChanged) {
    updates.push({
      range:
        `${safeSheetName}!F${existing.row}`,
      values: [[today]]
    });
  }

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId:
      MASTER_ROSTER_SPREADSHEET_ID,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: updates
    }
  });

  console.log(
    "MASTER ROSTER MEMBER UPDATED:",
    {
      discordId,
      row: existing.row,
      rankChanged,
      regiment,
      kompanie,
      position,
      active
    }
  );

  return {
    row: existing.row,
    created: false,
    rankChanged
  };
}


async function markMasterRosterMemberInactive({
  discordId,
  robloxUsername = null,
  rank = null
}) {
  const existing =
    await findMasterRosterMember(
      discordId
    );

  if (!existing) {
    console.warn(
      "MASTER ROSTER: member not found while marking inactive:",
      discordId
    );
    return null;
  }

  const sheetName =
    await getMasterRosterSheetName();

  const safeSheetName =
    escapeSheetName(sheetName);

  const data = [
    {
      range:
        `${safeSheetName}!J${existing.row}`,
      values: [["No"]]
    }
  ];

  if (robloxUsername) {
    data.push({
      range:
        `${safeSheetName}!B${existing.row}`,
      values: [[robloxUsername]]
    });
  }

  /*
   * If /removemember somehow sees a different rank than the Master Roster,
   * preserve the final rank and reset the current-rank date.
   */
  if (
    rank &&
    normalizeText(existing.rank) !==
      normalizeText(rank)
  ) {
    data.push(
      {
        range:
          `${safeSheetName}!D${existing.row}`,
        values: [[rank]]
      },
      {
        range:
          `${safeSheetName}!F${existing.row}`,
        values: [[
          getCurrentOrbatDate()
        ]]
      }
    );
  }

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId:
      MASTER_ROSTER_SPREADSHEET_ID,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data
    }
  });

  console.log(
    "MASTER ROSTER MEMBER MARKED INACTIVE:",
    {
      discordId,
      row: existing.row
    }
  );

  return existing.row;
}




function formatDiscordJoinDate(joinedAt) {
  if (!joinedAt) {
    return "";
  }

  const date =
    joinedAt instanceof Date
      ? joinedAt
      : new Date(joinedAt);

  if (
    Number.isNaN(
      date.getTime()
    )
  ) {
    return "";
  }

  const month =
    String(
      date.getMonth() + 1
    ).padStart(2, "0");

  const day =
    String(
      date.getDate()
    ).padStart(2, "0");

  const year =
    date.getFullYear();

  return `${month}/${day}/${year}`;
}


async function getDiscordArmyJoinDates(discordIds) {
  const joinDates =
    new Map();

  const uniqueIds =
    [
      ...new Set(
        discordIds
          .map(
            id =>
              String(id || "").trim()
          )
          .filter(Boolean)
      )
    ];

  if (
    uniqueIds.length === 0
  ) {
    return joinDates;
  }

  const guilds =
    [
      ...client.guilds.cache.values()
    ];

  /*
   * Fetch ONLY the Discord IDs found in the ORBAT.
   * This avoids guild.members.fetch() for the entire server, which
   * previously timed out with GuildMembersTimeout.
   *
   * We use small batches so Discord is not asked for the whole Army
   * in one request. If a batch request fails, fall back to individual
   * member fetches for that batch.
   */
  const MEMBER_BATCH_SIZE = 50;

  for (const guild of guilds) {
    for (
      let index = 0;
      index < uniqueIds.length;
      index += MEMBER_BATCH_SIZE
    ) {
      const batch =
        uniqueIds.slice(
          index,
          index + MEMBER_BATCH_SIZE
        );

      try {
        const members =
          await guild.members.fetch({
            user: batch,
            force: true
          });

        for (
          const member of
          members.values()
        ) {
          if (
            !member?.id ||
            !member.joinedAt
          ) {
            continue;
          }

          const existing =
            joinDates.get(
              member.id
            );

          if (
            !existing ||
            member.joinedAt.getTime() <
              existing.getTime()
          ) {
            joinDates.set(
              member.id,
              member.joinedAt
            );
          }
        }
      } catch (batchError) {
        console.warn(
          `[MASTER ROSTER] Batch Discord member fetch failed for guild ${guild.id}; falling back to individual member fetches.`
        );
        console.warn(
          batchError?.message ||
          batchError
        );

        for (const discordId of batch) {
          try {
            const member =
              await guild.members.fetch({
                user:
                  discordId,
                force:
                  true
              });

            if (
              member?.joinedAt
            ) {
              const existing =
                joinDates.get(
                  discordId
                );

              if (
                !existing ||
                member.joinedAt.getTime() <
                  existing.getTime()
              ) {
                joinDates.set(
                  discordId,
                  member.joinedAt
                );
              }
            }
          } catch (memberError) {
            /*
             * This normally means the Discord ID is not currently
             * a member of this guild, or Discord could not return it.
             * Do not fail the whole Master Roster sync.
             */
            console.warn(
              `[MASTER ROSTER] Could not fetch Discord member ${discordId} from guild ${guild.id}: ${memberError?.message || memberError}`
            );
          }
        }
      }
    }
  }

  return joinDates;
}


async function syncAllOrbatsToMasterRoster() {
  const result = {
    found: 0,
    uniqueDiscordIds: 0,
    added: 0,
    updated: 0,
    duplicates: 0,
    skipped: 0,
    errors: []
  };


  /*
   * STEP 1:
   * Read the Master Roster ONCE and build an in-memory lookup.
   */
  const masterSheetName =
    await getMasterRosterSheetName();

  const safeMasterSheetName =
    escapeSheetName(masterSheetName);

  const masterResponse =
    await sheets.spreadsheets.values.get({
      spreadsheetId:
        MASTER_ROSTER_SPREADSHEET_ID,
      range:
        `${safeMasterSheetName}!B${MASTER_ROSTER_FIRST_MEMBER_ROW}:J`,
      majorDimension: "ROWS"
    });

  const masterRows =
    masterResponse.data.values || [];

  const masterByDiscordId =
    new Map();

  let nextMasterRow =
    MASTER_ROSTER_FIRST_MEMBER_ROW;

  for (
    let index = 0;
    index < masterRows.length;
    index += 1
  ) {
    const rowNumber =
      MASTER_ROSTER_FIRST_MEMBER_ROW +
      index;

    const row =
      masterRows[index] || [];

    const discordId =
      String(row[1] || "").trim();

    if (discordId) {
      masterByDiscordId.set(
        discordId,
        {
          row: rowNumber,
          robloxUsername:
            String(row[0] || "").trim(),
          discordId,
          rank:
            String(row[2] || "").trim(),
          dateJoined:
            String(row[3] || "").trim(),
          dateOfCurrentRank:
            String(row[4] || "").trim(),
          regiment:
            String(row[5] || "").trim(),
          kompanie:
            String(row[6] || "").trim(),
          position:
            String(row[7] || "").trim(),
          active:
            String(row[8] || "").trim()
        }
      );
    }

    /*
     * Reuse the first completely empty Discord-ID row if there is one.
     * Otherwise append after the current roster data.
     */
    if (
      !discordId &&
      nextMasterRow ===
        MASTER_ROSTER_FIRST_MEMBER_ROW
    ) {
      nextMasterRow =
        rowNumber;
    }
  }

  if (
    masterRows.length > 0 &&
    nextMasterRow ===
      MASTER_ROSTER_FIRST_MEMBER_ROW &&
    String(
      masterRows[0]?.[1] || ""
    ).trim()
  ) {
    nextMasterRow =
      MASTER_ROSTER_FIRST_MEMBER_ROW +
      masterRows.length;
  }

  /*
   * STEP 2:
   * Read all ORBAT companies and collect members in memory.
   */
  const armyMembers =
    new Map();

  for (const regiment of REGIMENTS) {
    let companies;

    try {
      companies =
        await getCompanySheetNames(
          regiment
        );
    } catch (error) {
      result.errors.push(
        `${regiment.displayName}: ${error?.message || "Could not read company list."}`
      );
      continue;
    }

    for (const company of companies) {
      let summary;

      try {
        summary =
          await getCompanyRosterSummary({
            spreadsheetId:
              regiment.spreadsheetId,
            sheetName:
              company
          });
      } catch (error) {
        result.errors.push(
          `${regiment.displayName} / ${company}: ${error?.message || "Could not read company."}`
        );
        continue;
      }

      for (const member of summary.members) {
        result.found += 1;

        const discordId =
          String(
            member.discordId || ""
          ).trim();

        if (!discordId) {
          result.skipped += 1;
          continue;
        }

        if (
          armyMembers.has(
            discordId
          )
        ) {
          result.duplicates += 1;
          continue;
        }

        armyMembers.set(
          discordId,
          {
            discordId,
            robloxUsername:
              String(
                member.robloxUsername || ""
              ).trim(),
            rank:
              String(
                member.rank || ""
              ).trim(),
            regiment:
              regiment.displayName,
            kompanie:
              company,
            position:
              getMasterRosterPosition({
                regiment,
                company,
                row:
                  member.row
              })
          }
        );
      }
    }
  }

  result.uniqueDiscordIds =
    armyMembers.size;

  /*
   * STEP 3:
   * Now that the ORBAT scan is complete, fetch ONLY the Discord
   * members whose IDs actually appear in the Army ORBAT.
   */
  console.log(
    `[MASTER ROSTER] Fetching Discord join dates for ${armyMembers.size} ORBAT members...`
  );

  const discordJoinDates =
    await getDiscordArmyJoinDates(
      [
        ...armyMembers.keys()
      ]
    );

  console.log(
    `[MASTER ROSTER] Discord join dates loaded: ${discordJoinDates.size}/${armyMembers.size}`
  );

  /*
   * STEP 4:
   * Build ALL Master Roster writes in memory.
   *
   * Existing members preserve Date Joined Army.
   * Date of Current Rank changes only if the rank changed.
   */
  const today =
    getCurrentOrbatDate();

  const updates = [];

  for (
    const member of
    armyMembers.values()
  ) {
    const existing =
      masterByDiscordId.get(
        member.discordId
      );

    if (existing) {
      const rankChanged =
        normalizeText(
          existing.rank
        ) !==
        normalizeText(
          member.rank
        );

      const discordJoinedDate =
        formatDiscordJoinDate(
          discordJoinDates.get(
            member.discordId
          )
        );

      /*
       * 09/19/2026 was the placeholder date used by the first
       * Master Roster import. Replace that placeholder (or a blank
       * value) with the Discord server join date when available.
       *
       * Any other existing Date Joined Army is treated as intentional
       * historical data and is preserved.
       */
      const existingJoinDate =
        String(
          existing.dateJoined || ""
        ).trim();

      const shouldRepairJoinDate =
        !existingJoinDate ||
        existingJoinDate ===
          "09/19/2026" ||
        existingJoinDate ===
          "9/19/2026";

      const dateJoinedArmy =
        shouldRepairJoinDate
          ? (
              discordJoinedDate ||
              existingJoinDate ||
              ""
            )
          : existingJoinDate;

      const rowValues = [
        member.robloxUsername,
        member.discordId,
        member.rank,
        dateJoinedArmy,
        rankChanged
          ? today
          : (
              existing.dateOfCurrentRank ||
              today
            ),
        member.regiment,
        member.kompanie,
        member.position,
        "Yes"
      ];

      updates.push({
        range:
          `${safeMasterSheetName}!B${existing.row}:J${existing.row}`,
        values: [
          rowValues
        ]
      });

      result.updated += 1;
      continue;
    }

    /*
     * Find the next free row from our local snapshot.
     * This avoids another Google read for every new member.
     */
    while (true) {
      const index =
        nextMasterRow -
        MASTER_ROSTER_FIRST_MEMBER_ROW;

      const snapshotRow =
        masterRows[index] || [];

      const snapshotDiscordId =
        String(
          snapshotRow[1] || ""
        ).trim();

      const alreadyAssigned =
        updates.some(
          update =>
            update.range ===
            `${safeMasterSheetName}!B${nextMasterRow}:J${nextMasterRow}`
        );

      if (
        !snapshotDiscordId &&
        !alreadyAssigned
      ) {
        break;
      }

      nextMasterRow += 1;
    }

    const discordJoinedDate =
      formatDiscordJoinDate(
        discordJoinDates.get(
          member.discordId
        )
      );

    updates.push({
      range:
        `${safeMasterSheetName}!B${nextMasterRow}:J${nextMasterRow}`,
      values: [[
        member.robloxUsername,
        member.discordId,
        member.rank,
        discordJoinedDate,
        today,
        member.regiment,
        member.kompanie,
        member.position,
        "Yes"
      ]]
    });

    masterByDiscordId.set(
      member.discordId,
      {
        row:
          nextMasterRow,
        ...member,
        dateJoined:
          discordJoinedDate,
        dateOfCurrentRank:
          today,
        active:
          "Yes"
      }
    );

    result.added += 1;
    nextMasterRow += 1;
  }

  /*
   * STEP 5:
   * Batch write instead of member-by-member writes.
   * Google accepts many ranges in one values.batchUpdate request.
   */
  const BATCH_SIZE = 200;

  for (
    let index = 0;
    index < updates.length;
    index += BATCH_SIZE
  ) {
    const batch =
      updates.slice(
        index,
        index + BATCH_SIZE
      );

    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId:
        MASTER_ROSTER_SPREADSHEET_ID,
      requestBody: {
        valueInputOption:
          "USER_ENTERED",
        data:
          batch
      }
    });
  }

  console.log(
    "[MASTER ROSTER BATCH SYNC]",
    {
      found:
        result.found,
      uniqueDiscordIds:
        result.uniqueDiscordIds,
      added:
        result.added,
      updated:
        result.updated,
      duplicates:
        result.duplicates,
      skipped:
        result.skipped,
      errors:
        result.errors.length
    }
  );

  return result;
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


/*
|--------------------------------------------------------------------------
| Strike system
|--------------------------------------------------------------------------
|
| ORBAT Column F = current strike total (0-10) except Krümper/Garnison, where strikes are not displayed.
|
| Strike History is stored in the Army Master Roster spreadsheet so that
| reasons, issuers, removals, and automatic weekly reductions survive
| transfers between regiments/companies.
|
| Current agreed rules:
| - minimum issue: 1
| - normal AWOL guideline: 2
| - AWOL on own hosted event guideline: 5
| - maximum current total: 10
| - current total reduces by 1 for each full week elapsed
|
| /strike add accepts a free-text reason.
| /removestrike also requires a free-text reason.
|--------------------------------------------------------------------------
*/

const STRIKE_HISTORY_SHEET_NAME =
  "Strike History";

const STRIKE_STATE_SHEET_NAME =
  "Strike State";

const MAX_STRIKES = 10;

function getStrikeNowIso() {
  return new Date().toISOString();
}

function formatStrikeDate(value) {
  const date = new Date(value);

  if (Number.isNaN(date.getTime())) {
    return String(value || "Unknown");
  }

  return new Intl.DateTimeFormat(
    "en-US",
    {
      timeZone: "America/Los_Angeles",
      month: "2-digit",
      day: "2-digit",
      year: "numeric"
    }
  ).format(date);
}

async function ensureStrikeSheets() {
  const metadata =
    await sheets.spreadsheets.get({
      spreadsheetId:
        MASTER_ROSTER_SPREADSHEET_ID,
      fields:
        "sheets.properties.title"
    });

  const titles =
    new Set(
      (metadata.data.sheets || [])
        .map(
          sheet =>
            String(
              sheet.properties?.title ||
              ""
            )
        )
    );

  const requests = [];

  if (
    !titles.has(
      STRIKE_HISTORY_SHEET_NAME
    )
  ) {
    requests.push({
      addSheet: {
        properties: {
          title:
            STRIKE_HISTORY_SHEET_NAME
        }
      }
    });
  }

  if (
    !titles.has(
      STRIKE_STATE_SHEET_NAME
    )
  ) {
    requests.push({
      addSheet: {
        properties: {
          title:
            STRIKE_STATE_SHEET_NAME
        }
      }
    });
  }

  if (requests.length > 0) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId:
        MASTER_ROSTER_SPREADSHEET_ID,
      requestBody: {
        requests
      }
    });
  }

  const writes = [];

  if (
    !titles.has(
      STRIKE_HISTORY_SHEET_NAME
    )
  ) {
    writes.push({
      range:
        `'${STRIKE_HISTORY_SHEET_NAME}'!A1:J1`,
      values: [[
        "Timestamp",
        "Discord ID",
        "Roblox Username",
        "Type",
        "Change",
        "Reason",
        "Staff Discord ID",
        "Staff",
        "Regiment",
        "Kompanie"
      ]]
    });
  }

  if (
    !titles.has(
      STRIKE_STATE_SHEET_NAME
    )
  ) {
    writes.push({
      range:
        `'${STRIKE_STATE_SHEET_NAME}'!A1:B1`,
      values: [[
        "Discord ID",
        "Last Weekly Reduction Check"
      ]]
    });
  }

  if (writes.length > 0) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId:
        MASTER_ROSTER_SPREADSHEET_ID,
      requestBody: {
        valueInputOption:
          "RAW",
        data:
          writes
      }
    });
  }
}

async function getStrikeHistory(
  discordId
) {
  await ensureStrikeSheets();

  const response =
    await sheets.spreadsheets.values.get({
      spreadsheetId:
        MASTER_ROSTER_SPREADSHEET_ID,
      range:
        `'${STRIKE_HISTORY_SHEET_NAME}'!A2:J`,
      majorDimension:
        "ROWS"
    });

  return (
    response.data.values || []
  )
    .filter(
      row =>
        String(
          row?.[1] || ""
        ).trim() ===
        String(discordId).trim()
    )
    .map(row => ({
      timestamp:
        String(row?.[0] || ""),
      discordId:
        String(row?.[1] || ""),
      robloxUsername:
        String(row?.[2] || ""),
      type:
        String(row?.[3] || ""),
      change:
        Number.parseInt(
          String(row?.[4] || "0"),
          10
        ) || 0,
      reason:
        String(row?.[5] || ""),
      staffDiscordId:
        String(row?.[6] || ""),
      staff:
        String(row?.[7] || ""),
      regiment:
        String(row?.[8] || ""),
      kompanie:
        String(row?.[9] || "")
    }));
}

function calculateStrikeTotal(history) {
  return Math.max(
    0,
    Math.min(
      MAX_STRIKES,
      history.reduce(
        (total, entry) =>
          total + entry.change,
        0
      )
    )
  );
}

async function appendStrikeHistory({
  discordId,
  robloxUsername,
  type,
  change,
  reason,
  staffDiscordId,
  staff,
  regiment,
  kompanie
}) {
  await ensureStrikeSheets();

  await sheets.spreadsheets.values.append({
    spreadsheetId:
      MASTER_ROSTER_SPREADSHEET_ID,
    range:
      `'${STRIKE_HISTORY_SHEET_NAME}'!A:J`,
    valueInputOption:
      "RAW",
    insertDataOption:
      "INSERT_ROWS",
    requestBody: {
      values: [[
        getStrikeNowIso(),
        discordId,
        robloxUsername || "",
        type,
        change,
        reason,
        staffDiscordId || "",
        staff || "",
        regiment || "",
        kompanie || ""
      ]]
    }
  });
}

async function getStrikeState(
  discordId
) {
  await ensureStrikeSheets();

  const response =
    await sheets.spreadsheets.values.get({
      spreadsheetId:
        MASTER_ROSTER_SPREADSHEET_ID,
      range:
        `'${STRIKE_STATE_SHEET_NAME}'!A2:B`,
      majorDimension:
        "ROWS"
    });

  const rows =
    response.data.values || [];

  for (
    let index = 0;
    index < rows.length;
    index += 1
  ) {
    if (
      String(
        rows[index]?.[0] || ""
      ).trim() ===
      String(discordId).trim()
    ) {
      return {
        row:
          index + 2,
        lastCheck:
          String(
            rows[index]?.[1] || ""
          ).trim()
      };
    }
  }

  return null;
}

async function setStrikeState(
  discordId,
  dateIso
) {
  const state =
    await getStrikeState(
      discordId
    );

  if (state) {
    await sheets.spreadsheets.values.update({
      spreadsheetId:
        MASTER_ROSTER_SPREADSHEET_ID,
      range:
        `'${STRIKE_STATE_SHEET_NAME}'!A${state.row}:B${state.row}`,
      valueInputOption:
        "RAW",
      requestBody: {
        values: [[
          discordId,
          dateIso
        ]]
      }
    });
    return;
  }

  await sheets.spreadsheets.values.append({
    spreadsheetId:
      MASTER_ROSTER_SPREADSHEET_ID,
    range:
      `'${STRIKE_STATE_SHEET_NAME}'!A:B`,
    valueInputOption:
      "RAW",
    insertDataOption:
      "INSERT_ROWS",
    requestBody: {
      values: [[
        discordId,
        dateIso
      ]]
    }
  });
}

async function writeMemberStrikeTotal({
  existingMember,
  total
}) {
  const strikeColumn =
    getStrikeColumn(
      existingMember.companyName,
      existingMember.regiment.spreadsheetId
    );

  /*
   * Krümper and Garnison do not display strikes on their ORBAT.
   * Strike history/current total remains in the Master Roster
   * and is still shown by /memberinfo.
   */
  if (!strikeColumn) {
    console.log(
      "ORBAT STRIKE DISPLAY SKIPPED:",
      {
        company:
          existingMember.companyName,
        discordId:
          existingMember.discordId,
        total
      }
    );
    return;
  }

  const safeSheetName =
    escapeSheetName(
      existingMember.companyName
    );

  await sheets.spreadsheets.values.update({
    spreadsheetId:
      existingMember.regiment.spreadsheetId,
    range:
      `${safeSheetName}!${strikeColumn}${existingMember.row}`,
    valueInputOption:
      "RAW",
    requestBody: {
      values: [[
        Math.max(
          0,
          Math.min(
            MAX_STRIKES,
            total
          )
        )
      ]]
    }
  });  await clearLegacyTimezoneNoteFromStrikeCell({
    spreadsheetId,
    sheetName,
    row
  });

}

async function applyWeeklyStrikeReduction({
  discordId,
  existingMember,
  memberRecord
}) {
  let history =
    await getStrikeHistory(
      discordId
    );

  let total =
    calculateStrikeTotal(
      history
    );

  let state =
    await getStrikeState(
      discordId
    );

  const now =
    new Date();

  if (!state?.lastCheck) {
    await setStrikeState(
      discordId,
      now.toISOString()
    );

    await writeMemberStrikeTotal({
      existingMember,
      total
    });

    return {
      total,
      history,
      reducedBy: 0
    };
  }

  const lastCheck =
    new Date(
      state.lastCheck
    );

  if (
    Number.isNaN(
      lastCheck.getTime()
    )
  ) {
    await setStrikeState(
      discordId,
      now.toISOString()
    );

    await writeMemberStrikeTotal({
      existingMember,
      total
    });

    return {
      total,
      history,
      reducedBy: 0
    };
  }

  const weekMs =
    7 * 24 * 60 * 60 * 1000;

  const fullWeeks =
    Math.floor(
      (
        now.getTime() -
        lastCheck.getTime()
      ) /
      weekMs
    );

  if (fullWeeks <= 0) {
    await writeMemberStrikeTotal({
      existingMember,
      total
    });

    return {
      total,
      history,
      reducedBy: 0
    };
  }

  const reducedBy =
    Math.min(
      fullWeeks,
      total
    );

  if (reducedBy > 0) {
    await appendStrikeHistory({
      discordId,
      robloxUsername:
        memberRecord.robloxUsername,
      type:
        "Weekly Reduction",
      change:
        -reducedBy,
      reason:
        `Automatic weekly reduction for ${fullWeeks} full week(s).`,
      staffDiscordId:
        "SYSTEM",
      staff:
        "Prussian ORBAT Bot",
      regiment:
        existingMember.regiment.displayName,
      kompanie:
        existingMember.companyName
    });

    history =
      await getStrikeHistory(
        discordId
      );

    total =
      calculateStrikeTotal(
        history
      );
  }

  /*
   * Advance by the full number of elapsed weeks instead of resetting
   * the clock to an arbitrary partial-week boundary.
   */
  const advanced =
    new Date(
      lastCheck.getTime() +
      fullWeeks * weekMs
    );

  await setStrikeState(
    discordId,
    advanced.toISOString()
  );

  await writeMemberStrikeTotal({
    existingMember,
    total
  });

  return {
    total,
    history,
    reducedBy
  };
}

function buildStrikeHistoryLines(
  history,
  limit = 10
) {
  if (!history.length) {
    return [
      "No strike history."
    ];
  }

  return history
    .slice()
    .reverse()
    .slice(0, limit)
    .map(entry => {
      const sign =
        entry.change > 0
          ? "+"
          : "";

      const reason =
        entry.reason ||
        "No reason provided.";

      const staff =
        entry.staff ||
        entry.staffDiscordId ||
        "Unknown";

      return (
        `• **${sign}${entry.change} — ${entry.type}** ` +
        `(${formatStrikeDate(entry.timestamp)}) — ` +
        `${reason} — **By:** ${staff}`
      );
    });
}

async function getMemberRecord({
  spreadsheetId,
  sheetName,
  row
}) {
  const safeSheetName =
    escapeSheetName(sheetName);

  const {
    timezoneColumn,
    storageColumn
  } = getTimezoneColumnLayout(
    sheetName,
    spreadsheetId
  );

  const strikeColumn =
    getStrikeColumn(
      sheetName,
      spreadsheetId
    );

  const ranges = [
    `${safeSheetName}!C${row}:E${row}`,
    `${safeSheetName}!${timezoneColumn}${row}`,
    `${safeSheetName}!${storageColumn}${row}`
  ];

  if (strikeColumn) {
    ranges.push(
      `${safeSheetName}!${strikeColumn}${row}`
    );
  }

  const response =
    await sheets.spreadsheets.values.batchGet({
      spreadsheetId,
      ranges
    });

  const identity =
    response.data.valueRanges?.[0]?.values?.[0] || [];

  const timezoneValue =
    response.data.valueRanges?.[1]?.values?.[0]?.[0] || "";

  const timezoneStorageValue =
    response.data.valueRanges?.[3]?.values?.[0]?.[0] || "";

  const discordId =
    String(identity[1] || "").trim();

  let strikes = 0;

  if (strikeColumn) {
    strikes =
      Math.max(
        0,
        Math.min(
          MAX_STRIKES,
          Number.parseInt(
            String(
              response.data.valueRanges?.[2]?.values?.[0]?.[0] ||
              "0"
            ),
            10
          ) || 0
        )
      );
  } else if (discordId) {
    /*
     * Hidden-strike sheets get the authoritative total from
     * Master Roster Strike History instead of reading column F.
     */
    const history =
      await getStrikeHistory(
        discordId
      );

    strikes =
      calculateStrikeTotal(
        history
      );
  }

  return {
    robloxUsername:
      String(identity[0] || "").trim(),
    discordId,
    rank:
      String(identity[2] || "").trim(),
    strikes,
    timezone:
      String(timezoneValue || "").trim(),
    timezoneStorage:
      String(timezoneStorageValue || "").trim()
  };
}

async function getStandardAttendanceRecord({
  spreadsheetId,
  sheetName,
  row
}) {
  if (
    isAttendanceExcludedCompany(
      sheetName
    ) ||
    isSecondKrumperCompany(
      sheetName
    )
  ) {
    return null;
  }

  const safeSheetName =
    escapeSheetName(sheetName);

  const response =
    await sheets.spreadsheets.values.get({
      spreadsheetId,
      range:
        `${safeSheetName}!I${row}:O${row}`
    });

  const values =
    response.data.values?.[0] || [];

  return Array.from(
    { length: 7 },
    (_, index) =>
      values[index] ?? ""
  );
}


async function writeStandardAttendanceRecord({
  spreadsheetId,
  sheetName,
  row,
  attendance
}) {
  if (!Array.isArray(attendance)) {
    return;
  }

  const safeSheetName =
    escapeSheetName(sheetName);

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range:
      `${safeSheetName}!I${row}:O${row}`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[
        ...Array.from(
          { length: 7 },
          (_, index) =>
            attendance[index] ?? ""
        )
      ]]
    }
  });
}


async function removeMemberFromSheet({
  spreadsheetId,
  sheetName,
  row
}) {
  const safeSheetName =
    escapeSheetName(sheetName);

  const {
    timezoneColumn,
    storageColumn
  } = getTimezoneColumnLayout(
    sheetName,
    spreadsheetId
  );

  const strikeColumn =
    getStrikeColumn(
      sheetName,
      spreadsheetId
    );

  const ranges =
    isGarnisonCompany(sheetName)
      ? [
          `${safeSheetName}!C${row}:O${row}`
        ]
      : [
          `${safeSheetName}!C${row}`,
          `${safeSheetName}!D${row}`,
          `${safeSheetName}!E${row}`,
          ...(strikeColumn
            ? [`${safeSheetName}!${strikeColumn}${row}`]
            : []),
          `${safeSheetName}!${timezoneColumn}${row}`,
          `${safeSheetName}!${storageColumn}${row}`
        ];

  if (!isGarnisonCompany(sheetName)) {
    if (isFirstKrumperCompany(sheetName)) {
      ranges.push(
        `${safeSheetName}!I${row}`
      );
    } else if (
      isSecondKrumperCompany(sheetName)
    ) {
      ranges.push(
        `${safeSheetName}!I${row}:J${row}`
      );
    } else if (
      !isAttendanceExcludedCompany(sheetName)
    ) {
      ranges.push(
        `${safeSheetName}!I${row}:O${row}`
      );
    }
  }

  await sheets.spreadsheets.values.batchClear({
    spreadsheetId,
    requestBody: {
      ranges: [...new Set(ranges)]
    }
  });

  console.log(
    "ORBAT ROW CLEARED:",
    {
      sheetName,
      row,
      strikeColumn,
      timezoneColumn,
      storageColumn,
      clearedRanges:
        [...new Set(ranges)]
    }
  );
}

async function writeGarnisonInactivity({
  spreadsheetId,
  sheetName,
  row,
  duration,
  reason
}) {
  if (!isGarnisonCompany(sheetName)) {
    throw new Error(
      "GARNISON_INACTIVITY_WRONG_SHEET"
    );
  }

  const safeSheetName =
    escapeSheetName(sheetName);

  /*
   * Garnison merged fields:
   * J:M = Duration of Inactivity -> write to J, the top-left anchor.
   * N:O = Reason for Inactivity -> write to N, the top-left anchor.
   */
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: [
        {
          range: `${safeSheetName}!J${row}`,
          values: [[duration]]
        },
        {
          range: `${safeSheetName}!N${row}`,
          values: [[reason]]
        }
      ]
    }
  });
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
  strikes = 0,
  position = null,
  platoon = null
}) {
  const schuetzen =
    usesSchuetzenPositionStructure(
      spreadsheetId,
      sheetName
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

  const safeSheetName =
    escapeSheetName(sheetName);

  const {
    timezoneColumn,
    storageColumn
  } = getTimezoneColumnLayout(
    sheetName,
    spreadsheetId
  );

  const strikeColumn =
    getStrikeColumn(
      sheetName,
      spreadsheetId
    );

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
      range: `${safeSheetName}!${timezoneColumn}${row}`,
      values: [[timezone]]
    },
    {
      range: `${safeSheetName}!${storageColumn}${row}`,
      values: [[""]]
    }
  ];

  if (strikeColumn) {
    writeData.push({
      range:
        `${safeSheetName}!${strikeColumn}${row}`,
      values: [[
        Math.max(
          0,
          Math.min(
            MAX_STRIKES,
            Number.parseInt(
              String(strikes || "0"),
              10
            ) || 0
          )
        )
      ]]
    });
  }

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
      range: `${safeSheetName}!I${row}`,
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
  } else if (
    !schuetzen &&
    isGarnisonCompany(sheetName)
  ) {
    const dateAdded =
      getCurrentOrbatDate();

    writeData.push(
      {
        range: `${safeSheetName}!I${row}`,
        values: [[dateAdded]]
      },
      {
        /*
         * I is the top-left writable cell of merged I:L.
         */
        range: `${safeSheetName}!J${row}`,
        values: [[""]]
      },
      {
        /*
         * M is the top-left writable cell of merged M:N.
         */
        range: `${safeSheetName}!N${row}`,
        values: [[""]]
      }
    );

    console.log(
      "GARNISON DATE ADDED QUEUED:",
      {
        sheetName,
        row,
        date: dateAdded
      }
    );
  } else if (!schuetzen && !isAttendanceExcludedCompany(sheetName)) {
    /*
     * A Krümper entry date must NEVER follow a member into another company.
     * Explicitly blank H for ordinary non-Garnison, non-1. Krümper writes.
     */
    writeData.push({
      range: `${safeSheetName}!I${row}`,
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
| Final ORBAT column enforcement
|--------------------------------------------------------------------------
|
| This is intentionally run AFTER Apps Script timezone processing.
| It prevents an older/stale Apps Script deployment from moving a normal
| company's timezone back into column F.
|
| Krümper/Garnison:
|   F = timezone, G = IANA storage
|
| All other ORBAT sheets:
|   F = strikes, G = timezone, H = IANA storage
|--------------------------------------------------------------------------
*/

async function enforceOrbatTimezoneLayout({
  spreadsheetId,
  sheetName,
  row,
  timezone,
  ianaTimezone = ""
}) {
  const safeSheetName = escapeSheetName(sheetName);
  const layout = getOrbatColumnLayout(sheetName, spreadsheetId);

  /*
   * Apps Script returns storage values as:
   *   IANA:America/Los_Angeles
   *   OFFSET:-08:00
   *
   * Keep those prefixes because refreshAllTimezones() relies on them.
   * If an unprefixed IANA timezone reaches this safeguard, normalize it.
   */
  const rawStorageValue =
    String(ianaTimezone || "").trim();

  const timezoneStorageValue =
    rawStorageValue === ""
      ? ""
      : (
          rawStorageValue.startsWith("IANA:") ||
          rawStorageValue.startsWith("OFFSET:")
        )
          ? rawStorageValue
          : rawStorageValue.includes("/")
            ? `IANA:${rawStorageValue}`
            : rawStorageValue;

  if (!layout.strikeColumn) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: [
          {
            range: `${safeSheetName}!${layout.timezoneColumn}${row}`,
            values: [[timezone]]
          },
          {
            range: `${safeSheetName}!${layout.storageColumn}${row}`,
            values: [[timezoneStorageValue]]
          }
        ]
      }
    });

    console.log("[ORBAT LAYOUT ENFORCED]", {
      sheetName,
      row,
      strikes: "hidden",
      timezoneCell: `${layout.timezoneColumn}${row}`,
      storageCell: `${layout.storageColumn}${row}`
    });

    return;
  }

  // Preserve the existing strike total. If F was incorrectly overwritten
  // with a timezone by a stale Apps Script deployment, restore it to 0.
  const strikeRange = `${safeSheetName}!${layout.strikeColumn}${row}`;
  const strikeResponse = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: strikeRange
  });

  const currentStrikeValue = String(
    strikeResponse.data.values?.[0]?.[0] ?? ""
  ).trim();

  const parsedStrike = Number.parseInt(currentStrikeValue, 10);
  const safeStrike =
    Number.isFinite(parsedStrike)
      ? Math.max(0, Math.min(MAX_STRIKES, parsedStrike))
      : 0;

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: [
        {
          range: strikeRange,
          values: [[safeStrike]]
        },
        {
          range: `${safeSheetName}!${layout.timezoneColumn}${row}`,
          values: [[timezone]]
        },
        {
          range: `${safeSheetName}!${layout.storageColumn}${row}`,
          values: [[timezoneStorageValue]]
        }
      ]
    }
  });

  await clearLegacyTimezoneNoteFromStrikeCell({
    spreadsheetId,
    sheetName,
    row
  });

  console.log("[ORBAT LAYOUT ENFORCED]", {
    sheetName,
    row,
    strikeCell: `${layout.strikeColumn}${row}`,
    strikeValue: safeStrike,
    timezoneCell: `${layout.timezoneColumn}${row}`,
    storageCell: `${layout.storageColumn}${row}`
  });
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
        timezone,
        ...getTimezoneColumnLayout(sheetName, spreadsheetId)
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
| Multi-day interactive attendance board
|--------------------------------------------------------------------------
|
| /attendance flow:
| 1. Regiment + company
| 2. Choose one or more attendance days
| 3. Bot loads the actual ORBAT company roster
| 4. Edit one selected day at a time
| 5. Choose an attendance status
| 6. Select roster members for that status
| 7. Switch days, copy previous day, or mark remaining members AWOL
| 8. Review all selected days
| 9. Submit all changes to Google Sheets in one batch
|--------------------------------------------------------------------------
*/

const ATTENDANCE_SESSION_TTL_MS =
  20 * 60 * 1000;

const attendanceSessions =
  new Map();

const ATTENDANCE_STATUS_META = new Map([
  ["PRES", { label: "Present" }],
  ["EXC", { label: "Excused" }],
  ["AWOL", { label: "AWOL" }],
  ["RSVP", { label: "RSVP" }],
  ["MAYB", { label: "Maybe" }],
  ["NO", { label: "No" }],
  ["DM", { label: "DM" }],
  ["LEFT", { label: "Left" }]
]);

function getAttendanceStatusLabel(status) {
  const meta =
    ATTENDANCE_STATUS_META.get(
      String(status || "").trim()
    );

  return meta
    ? meta.label
    : String(status || "Unmarked");
}

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
  const session =
    attendanceSessions.get(token);

  if (!session) {
    return;
  }

  const delay =
    Math.max(
      1000,
      session.expiresAt - Date.now() + 1000
    );

  setTimeout(() => {
    const current =
      attendanceSessions.get(token);

    if (!current) {
      return;
    }

    if (
      current.expiresAt <= Date.now()
    ) {
      attendanceSessions.delete(token);
      return;
    }

    scheduleAttendanceSessionExpiry(token);
  }, delay);
}

async function ensureAttendanceRoster(
  session
) {
  if (
    Array.isArray(session.roster) &&
    session.roster.length > 0
  ) {
    return session.roster;
  }

  const summary =
    await getCompanyRosterSummary({
      spreadsheetId:
        session.spreadsheetId,
      sheetName:
        session.company
    });

  session.roster =
    summary.members
      .filter(member =>
        member.robloxUsername ||
        member.discordId ||
        member.rank
      )
      .map(member => ({
        row: member.row,
        robloxUsername:
          member.robloxUsername ||
          `Row ${member.row}`,
        discordId:
          member.discordId || "",
        rank:
          member.rank || "No rank"
      }));

  return session.roster;
}

function getAvailableAttendanceDays(
  company
) {
  return isSecondKrumperCompany(company)
    ? ["Saturday", "Sunday"]
    : ATTENDANCE_DAYS;
}

function getDayPendingMap(
  session,
  day = session.activeDay
) {
  if (!(session.pendingByDay instanceof Map)) {
    session.pendingByDay = new Map();
  }

  if (!day) {
    return new Map();
  }

  if (!session.pendingByDay.has(day)) {
    session.pendingByDay.set(
      day,
      new Map()
    );
  }

  return session.pendingByDay.get(day);
}

function getAttendanceCountsForDay(
  session,
  day = session.activeDay
) {
  const pending =
    getDayPendingMap(
      session,
      day
    );

  const counts =
    new Map();

  for (
    const status of
    ATTENDANCE_STATUS_CHOICES
  ) {
    counts.set(status, 0);
  }

  for (const status of pending.values()) {
    counts.set(
      status,
      (counts.get(status) || 0) + 1
    );
  }

  return counts;
}

function getAttendanceDayIndex(
  session,
  day = session.activeDay
) {
  return Array.isArray(session.days)
    ? session.days.indexOf(day)
    : -1;
}

function getPreviousAttendanceDay(
  session
) {
  const index =
    getAttendanceDayIndex(session);

  return index > 0
    ? session.days[index - 1]
    : null;
}

function getNextAttendanceDay(
  session
) {
  const index =
    getAttendanceDayIndex(session);

  return (
    index >= 0 &&
    index < session.days.length - 1
  )
    ? session.days[index + 1]
    : null;
}

function attendanceSessionHeader(
  session
) {
  const rosterCount =
    Array.isArray(session.roster)
      ? session.roster.length
      : 0;

  const selectedDays =
    Array.isArray(session.days) &&
    session.days.length > 0
      ? session.days.join(", ")
      : "Not selected";

  return [
    "**Multi-Day Company Attendance Board**",
    "",
    `**Regiment:** ${session.regimentDisplayName}`,
    `**Company:** ${session.company}`,
    `**Selected Days:** ${selectedDays}`,
    session.activeDay
      ? `**Editing:** ${session.activeDay}`
      : null,
    rosterCount > 0
      ? `**Company Strength:** ${rosterCount}`
      : null
  ].filter(Boolean);
}

function attendanceBoardLines(
  session
) {
  const pending =
    getDayPendingMap(
      session
    );

  const counts =
    getAttendanceCountsForDay(
      session
    );

  const rosterCount =
    Array.isArray(session.roster)
      ? session.roster.length
      : 0;

  const unmarked =
    Math.max(
      0,
      rosterCount - pending.size
    );

  const dayIndex =
    getAttendanceDayIndex(session);

  const lines = [
    ...attendanceSessionHeader(session),
    "",
    `**Day ${dayIndex + 1} of ${session.days.length}**`,
    `**Unmarked:** ${unmarked}`,
    `**Present:** ${counts.get("PRES") || 0}`,
    `**Excused:** ${counts.get("EXC") || 0}`,
    `**AWOL:** ${counts.get("AWOL") || 0}`
  ];

  const otherStatuses =
    ATTENDANCE_STATUS_CHOICES
      .filter(status =>
        !["PRES", "EXC", "AWOL"].includes(
          status
        )
      )
      .filter(status =>
        (counts.get(status) || 0) > 0
      )
      .map(
        status =>
          `**${getAttendanceStatusLabel(status)}:** ` +
          `${counts.get(status)}`
      );

  if (otherStatuses.length > 0) {
    lines.push(...otherStatuses);
  }

  lines.push(
    "",
    "Choose a status, then select the roster members for that status.",
    "Switch days at any time. Nothing is written until **Submit All Attendance**."
  );

  return lines;
}

function buildAttendanceDayComponents(
  session
) {
  const availableDays =
    getAvailableAttendanceDays(
      session.company
    );

  const dayMenu =
    new StringSelectMenuBuilder()
      .setCustomId(
        `attendance_days:${session.token}`
      )
      .setPlaceholder(
        "Choose one or more attendance days"
      )
      .setMinValues(1)
      .setMaxValues(
        availableDays.length
      )
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

function buildAttendanceBoardComponents(
  session
) {
  const daySwitch =
    new StringSelectMenuBuilder()
      .setCustomId(
        `attendance_switch_day:${session.token}`
      )
      .setPlaceholder(
        `Editing ${session.activeDay}`
      )
      .setMinValues(1)
      .setMaxValues(1)
      .addOptions(
        session.days.map(day => ({
          label:
            day === session.activeDay
              ? `${day} (Current)`
              : day,
          value: day,
          default:
            day === session.activeDay
        }))
      );

  const statusOptions =
    ATTENDANCE_STATUS_CHOICES
      .map(status => {
        const meta =
          ATTENDANCE_STATUS_META.get(status);

        return {
          label:
            meta?.label || status,
          value: status,
          description:
            `Mark selected members as ${status}`
        };
      });

  statusOptions.push({
    label: "Unmark / Clear",
    value: "__UNMARK__",
    description:
      "Remove pending attendance from selected members"
  });

  const statusMenu =
    new StringSelectMenuBuilder()
      .setCustomId(
        `attendance_pick_status:${session.token}`
      )
      .setPlaceholder(
        "Choose a status to assign"
      )
      .setMinValues(1)
      .setMaxValues(1)
      .addOptions(statusOptions);

  const previous =
    new ButtonBuilder()
      .setCustomId(
        `attendance_previous_day:${session.token}`
      )
      .setLabel("Previous Day")
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(
        !getPreviousAttendanceDay(
          session
        )
      );

  const next =
    new ButtonBuilder()
      .setCustomId(
        `attendance_next_day:${session.token}`
      )
      .setLabel("Next Day")
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(
        !getNextAttendanceDay(
          session
        )
      );

  const copyPrevious =
    new ButtonBuilder()
      .setCustomId(
        `attendance_copy_previous:${session.token}`
      )
      .setLabel("Copy Previous Day")
      .setStyle(ButtonStyle.Primary)
      .setDisabled(
        !getPreviousAttendanceDay(
          session
        )
      );

  const remainingAwol =
    new ButtonBuilder()
      .setCustomId(
        `attendance_remaining_awol:${session.token}`
      )
      .setLabel("Remaining -> AWOL")
      .setStyle(ButtonStyle.Danger);

  const review =
    new ButtonBuilder()
      .setCustomId(
        `attendance_review:${session.token}`
      )
      .setLabel("Review All Days")
      .setStyle(ButtonStyle.Success);

  const clearDay =
    new ButtonBuilder()
      .setCustomId(
        `attendance_clear_day:${session.token}`
      )
      .setLabel("Clear Current Day")
      .setStyle(ButtonStyle.Secondary);

  const changeDays =
    new ButtonBuilder()
      .setCustomId(
        `attendance_change_days:${session.token}`
      )
      .setLabel("Change Selected Days")
      .setStyle(ButtonStyle.Secondary);

  const cancel =
    new ButtonBuilder()
      .setCustomId(
        `attendance_cancel:${session.token}`
      )
      .setLabel("Cancel")
      .setStyle(ButtonStyle.Danger);

  return [
    new ActionRowBuilder()
      .addComponents(daySwitch),
    new ActionRowBuilder()
      .addComponents(statusMenu),
    new ActionRowBuilder()
      .addComponents(
        previous,
        next,
        copyPrevious,
        remainingAwol,
        review
      ),
    new ActionRowBuilder()
      .addComponents(
        clearDay,
        changeDays,
        cancel
      )
  ];
}

function buildAttendanceRosterComponents(
  session
) {
  const roster =
    Array.isArray(session.roster)
      ? session.roster
      : [];

  const pending =
    getDayPendingMap(
      session
    );

  const options =
    roster
      .slice(0, 25)
      .map(member => {
        const current =
          pending.get(
            String(member.row)
          );

        const descriptionParts = [
          member.rank
        ];

        if (current) {
          descriptionParts.push(
            `Pending: ${getAttendanceStatusLabel(current)}`
          );
        } else {
          descriptionParts.push(
            "Pending: Unmarked"
          );
        }

        return {
          label:
            String(
              member.robloxUsername ||
              `Row ${member.row}`
            ).slice(0, 100),
          value:
            String(member.row),
          description:
            descriptionParts
              .join(" - ")
              .slice(0, 100)
        };
      });

  const rows = [];

  if (options.length > 0) {
    const memberMenu =
      new StringSelectMenuBuilder()
        .setCustomId(
          `attendance_roster:${session.token}`
        )
        .setPlaceholder(
          session.selectedStatus ===
            "__UNMARK__"
            ? `Select ${session.activeDay} members to unmark`
            : `Select ${session.activeDay} members for ${getAttendanceStatusLabel(session.selectedStatus)}`
        )
        .setMinValues(1)
        .setMaxValues(
          Math.min(
            25,
            options.length
          )
        )
        .addOptions(options);

    rows.push(
      new ActionRowBuilder()
        .addComponents(memberMenu)
    );
  }

  const back =
    new ButtonBuilder()
      .setCustomId(
        `attendance_back_board:${session.token}`
      )
      .setLabel("Back to Board")
      .setStyle(ButtonStyle.Secondary);

  const cancel =
    new ButtonBuilder()
      .setCustomId(
        `attendance_cancel:${session.token}`
      )
      .setLabel("Cancel")
      .setStyle(ButtonStyle.Danger);

  rows.push(
    new ActionRowBuilder()
      .addComponents(
        back,
        cancel
      )
  );

  return rows;
}

function attendanceReviewLines(
  session
) {
  const rosterCount =
    Array.isArray(session.roster)
      ? session.roster.length
      : 0;

  const lines = [
    "**Review Multi-Day Attendance**",
    "",
    `**Regiment:** ${session.regimentDisplayName}`,
    `**Company:** ${session.company}`,
    `**Company Strength:** ${rosterCount}`,
    `**Days:** ${session.days.join(", ")}`
  ];

  let totalPending = 0;

  for (const day of session.days) {
    const pending =
      getDayPendingMap(
        session,
        day
      );

    const counts =
      getAttendanceCountsForDay(
        session,
        day
      );

    totalPending +=
      pending.size;

    lines.push(
      "",
      `**${day}**`,
      `Present: ${counts.get("PRES") || 0} | ` +
      `Excused: ${counts.get("EXC") || 0} | ` +
      `AWOL: ${counts.get("AWOL") || 0} | ` +
      `Unmarked: ${Math.max(0, rosterCount - pending.size)}`
    );

    const extras =
      ATTENDANCE_STATUS_CHOICES
        .filter(status =>
          !["PRES", "EXC", "AWOL"].includes(
            status
          )
        )
        .filter(status =>
          (counts.get(status) || 0) > 0
        )
        .map(
          status =>
            `${getAttendanceStatusLabel(status)}: ${counts.get(status)}`
        );

    if (extras.length > 0) {
      lines.push(
        extras.join(" | ")
      );
    }
  }

  lines.push(
    "",
    `**Total Pending Entries:** ${totalPending}`,
    "",
    "Press **Submit All Attendance** to write every selected day to Google Sheets."
  );

  return lines;
}

function getTotalPendingAttendanceEntries(
  session
) {
  if (!(session.pendingByDay instanceof Map)) {
    return 0;
  }

  let total = 0;

  for (const day of session.days || []) {
    total +=
      getDayPendingMap(
        session,
        day
      ).size;
  }

  return total;
}

function buildAttendanceReviewComponents(
  session
) {
  const submit =
    new ButtonBuilder()
      .setCustomId(
        `attendance_submit_all:${session.token}`
      )
      .setLabel("Submit All Attendance")
      .setStyle(ButtonStyle.Success)
      .setDisabled(
        getTotalPendingAttendanceEntries(
          session
        ) === 0
      );

  const back =
    new ButtonBuilder()
      .setCustomId(
        `attendance_back_board:${session.token}`
      )
      .setLabel("Back to Board")
      .setStyle(ButtonStyle.Secondary);

  const cancel =
    new ButtonBuilder()
      .setCustomId(
        `attendance_cancel:${session.token}`
      )
      .setLabel("Cancel")
      .setStyle(ButtonStyle.Danger);

  return [
    new ActionRowBuilder()
      .addComponents(
        submit,
        back,
        cancel
      )
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
        "Attendance entry cancelled. No spreadsheet changes were made.",
      components: []
    });

    return true;
  }

  if (
    action ===
    "attendance_days"
  ) {
    const days =
      interaction.values || [];

    if (days.length === 0) {
      await interaction.reply({
        content:
          "Select at least one attendance day.",
        flags:
          MessageFlags.Ephemeral
      });

      return true;
    }

    for (const day of days) {
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
              : `The attendance day ${day} is not valid for this company.`,
          flags:
            MessageFlags.Ephemeral
        });

        return true;
      }
    }

    const orderedDays =
      getAvailableAttendanceDays(
        session.company
      ).filter(day =>
        days.includes(day)
      );

    session.days =
      orderedDays;

    session.activeDay =
      session.days[0];

    session.selectedStatus =
      null;

    session.pendingByDay =
      new Map();

    for (const day of session.days) {
      session.pendingByDay.set(
        day,
        new Map()
      );
    }

    await ensureAttendanceRoster(
      session
    );

    if (session.roster.length === 0) {
      await interaction.update({
        content: [
          ...attendanceSessionHeader(
            session
          ),
          "",
          "No members were found in this company roster."
        ].join("\n"),
        components: [
          new ActionRowBuilder()
            .addComponents(
              new ButtonBuilder()
                .setCustomId(
                  `attendance_cancel:${session.token}`
                )
                .setLabel("Close")
                .setStyle(
                  ButtonStyle.Secondary
                )
            )
        ]
      });

      return true;
    }

    await interaction.update({
      content:
        attendanceBoardLines(
          session
        ).join("\n"),
      components:
        buildAttendanceBoardComponents(
          session
        )
    });

    return true;
  }

  if (
    action ===
    "attendance_switch_day"
  ) {
    const day =
      String(
        interaction.values?.[0] || ""
      ).trim();

    if (
      !session.days.includes(day)
    ) {
      await interaction.reply({
        content:
          "That day is not part of this attendance session.",
        flags:
          MessageFlags.Ephemeral
      });

      return true;
    }

    session.activeDay =
      day;

    session.selectedStatus =
      null;

    await interaction.update({
      content:
        attendanceBoardLines(
          session
        ).join("\n"),
      components:
        buildAttendanceBoardComponents(
          session
        )
    });

    return true;
  }

  if (
    action ===
    "attendance_previous_day"
  ) {
    const previous =
      getPreviousAttendanceDay(
        session
      );

    if (!previous) {
      await interaction.reply({
        content:
          "There is no previous selected day.",
        flags:
          MessageFlags.Ephemeral
      });

      return true;
    }

    session.activeDay =
      previous;

    session.selectedStatus =
      null;

    await interaction.update({
      content:
        attendanceBoardLines(
          session
        ).join("\n"),
      components:
        buildAttendanceBoardComponents(
          session
        )
    });

    return true;
  }

  if (
    action ===
    "attendance_next_day"
  ) {
    const next =
      getNextAttendanceDay(
        session
      );

    if (!next) {
      await interaction.reply({
        content:
          "There is no next selected day.",
        flags:
          MessageFlags.Ephemeral
      });

      return true;
    }

    session.activeDay =
      next;

    session.selectedStatus =
      null;

    await interaction.update({
      content:
        attendanceBoardLines(
          session
        ).join("\n"),
      components:
        buildAttendanceBoardComponents(
          session
        )
    });

    return true;
  }

  if (
    action ===
    "attendance_copy_previous"
  ) {
    const previous =
      getPreviousAttendanceDay(
        session
      );

    if (!previous) {
      await interaction.reply({
        content:
          "There is no previous selected day to copy.",
        flags:
          MessageFlags.Ephemeral
      });

      return true;
    }

    const previousMap =
      getDayPendingMap(
        session,
        previous
      );

    session.pendingByDay.set(
      session.activeDay,
      new Map(previousMap)
    );

    await interaction.update({
      content: [
        ...attendanceBoardLines(
          session
        ),
        "",
        `Copied all pending attendance from **${previous}** into **${session.activeDay}**.`
      ].join("\n"),
      components:
        buildAttendanceBoardComponents(
          session
        )
    });

    return true;
  }

  if (
    action ===
    "attendance_pick_status"
  ) {
    const status =
      String(
        interaction.values?.[0] || ""
      ).trim();

    if (
      status !== "__UNMARK__" &&
      !ATTENDANCE_STATUS_CHOICES.includes(
        status
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

    session.selectedStatus =
      status;

    await ensureAttendanceRoster(
      session
    );

    await interaction.update({
      content: [
        ...attendanceSessionHeader(
          session
        ),
        "",
        status === "__UNMARK__"
          ? `Select the ${session.activeDay} roster members whose pending attendance should be cleared.`
          : `Select the ${session.activeDay} roster members to mark as **${getAttendanceStatusLabel(status)}**.`
      ].join("\n"),
      components:
        buildAttendanceRosterComponents(
          session
        )
    });

    return true;
  }

  if (
    action ===
    "attendance_roster"
  ) {
    const selectedRows =
      interaction.values || [];

    const pending =
      getDayPendingMap(
        session
      );

    let changed = 0;

    for (const row of selectedRows) {
      const key =
        String(row);

      if (
        session.selectedStatus ===
        "__UNMARK__"
      ) {
        if (pending.delete(key)) {
          changed += 1;
        }
      } else {
        pending.set(
          key,
          session.selectedStatus
        );
        changed += 1;
      }
    }

    session.selectedStatus = null;

    await interaction.update({
      content: [
        ...attendanceBoardLines(
          session
        ),
        "",
        `Updated **${changed}** member(s) for **${session.activeDay}**.`
      ].join("\n"),
      components:
        buildAttendanceBoardComponents(
          session
        )
    });

    return true;
  }

  if (
    action ===
    "attendance_remaining_awol"
  ) {
    await ensureAttendanceRoster(
      session
    );

    const pending =
      getDayPendingMap(
        session
      );

    let added = 0;

    for (const member of session.roster) {
      const key =
        String(member.row);

      if (!pending.has(key)) {
        pending.set(
          key,
          "AWOL"
        );
        added += 1;
      }
    }

    await interaction.update({
      content: [
        ...attendanceBoardLines(
          session
        ),
        "",
        `Marked **${added}** previously unmarked member(s) as **AWOL** for **${session.activeDay}**.`
      ].join("\n"),
      components:
        buildAttendanceBoardComponents(
          session
        )
    });

    return true;
  }

  if (
    action ===
    "attendance_clear_day"
  ) {
    getDayPendingMap(
      session
    ).clear();

    session.selectedStatus = null;

    await interaction.update({
      content: [
        ...attendanceBoardLines(
          session
        ),
        "",
        `Cleared all pending attendance for **${session.activeDay}**.`
      ].join("\n"),
      components:
        buildAttendanceBoardComponents(
          session
        )
    });

    return true;
  }

  if (
    action ===
    "attendance_change_days"
  ) {
    session.days = [];
    session.activeDay = null;
    session.selectedStatus = null;
    session.pendingByDay = new Map();

    await interaction.update({
      content: [
        ...attendanceSessionHeader(
          session
        ),
        "",
        "Choose the attendance days again.",
        "Changing selected days cleared all pending attendance marks."
      ].join("\n"),
      components:
        buildAttendanceDayComponents(
          session
        )
    });

    return true;
  }

  if (
    action ===
    "attendance_back_board"
  ) {
    session.selectedStatus = null;

    await interaction.update({
      content:
        attendanceBoardLines(
          session
        ).join("\n"),
      components:
        buildAttendanceBoardComponents(
          session
        )
    });

    return true;
  }

  if (
    action ===
    "attendance_review"
  ) {
    if (
      getTotalPendingAttendanceEntries(
        session
      ) === 0
    ) {
      await interaction.reply({
        content:
          "No attendance has been marked yet.",
        flags:
          MessageFlags.Ephemeral
      });

      return true;
    }

    await interaction.update({
      content:
        attendanceReviewLines(
          session
        ).join("\n"),
      components:
        buildAttendanceReviewComponents(
          session
        )
    });

    return true;
  }

  if (
    action ===
    "attendance_submit_all"
  ) {
    if (
      getTotalPendingAttendanceEntries(
        session
      ) === 0
    ) {
      await interaction.reply({
        content:
          "There are no pending attendance changes to submit.",
        flags:
          MessageFlags.Ephemeral
      });

      return true;
    }

    await interaction.deferUpdate();

    const safeSheetName =
      escapeSheetName(
        session.company
      );

    const rosterByRow =
      new Map(
        session.roster.map(
          member => [
            String(member.row),
            member
          ]
        )
      );

    const data = [];
    const submitted = [];

    for (const day of session.days) {
      const attendanceColumn =
        getAttendanceColumn(
          session.company,
          day
        );

      const pending =
        getDayPendingMap(
          session,
          day
        );

      for (
        const [rowKey, status] of
        pending.entries()
      ) {
        const member =
          rosterByRow.get(rowKey);

        if (!member) {
          continue;
        }

        data.push({
          range:
            `${safeSheetName}!${attendanceColumn}${member.row}`,
          values: [[status]]
        });

        submitted.push({
          ...member,
          day,
          status,
          cell:
            `${attendanceColumn}${member.row}`
        });
      }
    }

    if (data.length === 0) {
      attendanceSessions.delete(token);

      await interaction.editReply({
        content:
          "No valid roster attendance changes were available to submit.",
        components: []
      });

      return true;
    }

    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId:
        session.spreadsheetId,
      requestBody: {
        valueInputOption:
          "USER_ENTERED",
        data
      }
    });

    const summaryLines = [
      "**Multi-Day Attendance Submitted**",
      "",
      `**Regiment:** ${session.regimentDisplayName}`,
      `**Company:** ${session.company}`,
      `**Days Submitted:** ${session.days.join(", ")}`,
      `**Attendance Entries Updated:** ${submitted.length}`
    ];

    for (const day of session.days) {
      const counts =
        getAttendanceCountsForDay(
          session,
          day
        );

      const tracked =
        (counts.get("PRES") || 0) +
        (counts.get("EXC") || 0) +
        (counts.get("AWOL") || 0);

      const rate =
        tracked > 0
          ? (counts.get("PRES") || 0) /
            tracked
          : null;

      summaryLines.push(
        "",
        `**${day}**`,
        `Present: ${counts.get("PRES") || 0} | ` +
        `Excused: ${counts.get("EXC") || 0} | ` +
        `AWOL: ${counts.get("AWOL") || 0} | ` +
        `Attendance Rate: ${rate === null ? "N/A" : `${(rate * 100).toFixed(0)}%`}`
      );
    }

    summaryLines.push(
      "",
      `Recorded by <@${interaction.user.id}>`
    );

    await sendOrbatLog({
      interaction,
      category: "Attendance",
      action:
        "Multi-Day Company Attendance Submitted",
      affectedMember: null,
      robloxUsername:
        `${submitted.length} attendance entry/entries`,
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
          label: "Days",
          after:
            session.days.join(", ")
        },
        {
          label: "Entries Updated",
          after:
            String(
              submitted.length
            )
        }
      ],
      notes:
        submitted
          .slice(0, 20)
          .map(
            member =>
              `${member.day}: ${member.robloxUsername} -> ${member.status} (${member.cell})`
          )
          .join("\n")
    });

    attendanceSessions.delete(token);

    await interaction.editReply({
      content:
        summaryLines.join("\n"),
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
    console.log(
      "[INTERACTION RECEIVED]",
      {
        type: interaction.type,
        commandName:
          interaction.commandName || null,
        isChatInputCommand:
          interaction.isChatInputCommand(),
        isAutocomplete:
          interaction.isAutocomplete()
      }
    );
    if (interaction.isAutocomplete()) {
      const isCompanyAutocompleteCommand =
        interaction.commandName ===
          "addmember" ||
        interaction.commandName ===
          "transfer" ||
        interaction.commandName ===
          "attendance" ||
        interaction.commandName ===
          "organize" ||
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

    /*
     * ACKNOWLEDGE STRIKE COMMANDS IMMEDIATELY
     *
     * Discord only gives a slash command a short window to acknowledge
     * the interaction. /strike and /removestrike can perform several
     * Google Sheets / Master Roster operations, so defer them here,
     * before any other command routing or sheet work occurs.
     */
    if (
      interaction.commandName === "strike" ||
      interaction.commandName === "removestrike"
    ) {
      try {
        if (
          !interaction.deferred &&
          !interaction.replied
        ) {
          await interaction.deferReply({
            flags:
              MessageFlags.Ephemeral
          });
        }

        console.log(
          "[STRIKE COMMAND ACKNOWLEDGED]",
          {
            commandName:
              interaction.commandName,
            userId:
              interaction.user?.id || null
          }
        );
      } catch (ackError) {
        console.error(
          "[STRIKE COMMAND ACKNOWLEDGEMENT FAILED]",
          ackError
        );
        return;
      }
    }

    /*
     * Handle /syncmasterroster FIRST.
     *
     * This intentionally sits immediately after the ChatInputCommand guard
     * so no other command-routing logic can intercept the interaction before
     * Discord receives an acknowledgement.
     */
    if (
      interaction.commandName ===
      "syncmasterroster"
    ) {
      console.log(
        "[SYNC MASTER ROSTER] Handler entered"
      );

      try {
        await interaction.deferReply({
          flags:
            MessageFlags.Ephemeral
        });

        console.log(
          "[SYNC MASTER ROSTER] Reply deferred"
        );

        console.log(
          "[SYNC MASTER ROSTER] Starting ORBAT scan"
        );

        const result =
          await syncAllOrbatsToMasterRoster();

        console.log(
          "[SYNC MASTER ROSTER] Scan completed",
          result
        );

        const lines = [
          "**Army Master Roster Synchronization Complete**",
          "",
          `**ORBAT Members Found:** ${result.found}`,
          `**Unique Discord IDs:** ${result.uniqueDiscordIds}`,
          `**New Members Added:** ${result.added}`,
          `**Existing Members Updated:** ${result.updated}`,
          `**Duplicate Discord IDs Skipped:** ${result.duplicates}`,
          `**Members Without Discord IDs Skipped:** ${result.skipped}`,
          `**Errors:** ${result.errors.length}`
        ];

        if (
          result.errors.length > 0
        ) {
          lines.push(
            "",
            "**First Errors:**",
            ...result.errors
              .slice(0, 8)
              .map(
                error =>
                  `• ${error}`
              )
          );

          if (
            result.errors.length > 8
          ) {
            lines.push(
              `• …and ${result.errors.length - 8} more error(s). Check Railway logs for the full list.`
            );
          }

          console.error(
            "[SYNC MASTER ROSTER] ERRORS:",
            result.errors
          );
        }

        await interaction.editReply(
          lines
            .join("\n")
            .slice(0, 1950)
        );

        console.log(
          "[SYNC MASTER ROSTER] Discord response sent"
        );

        return;
      } catch (error) {
        console.error(
          "[SYNC MASTER ROSTER] FAILED:"
        );
        console.error(error);

        try {
          if (
            interaction.deferred ||
            interaction.replied
          ) {
            await interaction.editReply(
              [
                "The Army Master Roster synchronization failed.",
                "",
                `**Error:** ${error?.message || "Unknown error."}`,
                "",
                "No ORBAT members were removed by this command. Check the Railway logs for details."
              ].join("\n")
            );
          } else {
            await interaction.reply({
              content: [
                "The Army Master Roster synchronization failed.",
                "",
                `**Error:** ${error?.message || "Unknown error."}`
              ].join("\n"),
              flags:
                MessageFlags.Ephemeral
            });
          }
        } catch (replyError) {
          console.error(
            "[SYNC MASTER ROSTER] Failed to send error response:"
          );
          console.error(replyError);
        }

        return;
      }
    }

    if (
      interaction.commandName !== "addmember" &&
      interaction.commandName !== "removemember" &&
      interaction.commandName !== "rank" &&
      interaction.commandName !== "transfer" &&
      interaction.commandName !== "attendance" &&
      interaction.commandName !== "organize" &&
      interaction.commandName !== "memberinfo" &&
      interaction.commandName !== "companyinfo" &&
      interaction.commandName !== "missingattendance" &&
      interaction.commandName !== "getsheet" &&
      interaction.commandName !== "roster" &&
      interaction.commandName !== "strength" &&
      interaction.commandName !== "attendanceview" &&
      interaction.commandName !== "audit" &&
      interaction.commandName !== "strike" &&
      interaction.commandName !== "removestrike"
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

    if (interaction.commandName === "strike") {
      try {
        if (
          !interaction.deferred &&
          !interaction.replied
        ) {
          await interaction.deferReply({
            flags:
              MessageFlags.Ephemeral
          });
        }

        console.log(
          "[STRIKE ADD] Handler entered"
        );

        const subcommand =
          interaction.options.getSubcommand(
            true
          );

        if (subcommand !== "add") {
          await interaction.editReply(
            "Unknown strike action."
          );
          return;
        }

        const discordMember =
          interaction.options.getUser(
            "discord_member",
            true
          );

        const strikeType =
          interaction.options.getString(
            "type",
            true
          );

        const amount =
          interaction.options.getInteger(
            "amount",
            true
          );

        const reason =
          interaction.options.getString(
            "reason",
            true
          ).trim();

        const existingMember =
          await findMemberByDiscordId(
            discordMember.id
          );

        if (!existingMember) {
          await interaction.editReply(
            "That member was not found in the Grand ORBAT."
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

        const before =
          await applyWeeklyStrikeReduction({
            discordId:
              discordMember.id,
            existingMember,
            memberRecord
          });

        if (
          before.total + amount >
          MAX_STRIKES
        ) {
          await interaction.editReply(
            [
              "The strike was not added.",
              "",
              `**Current Strikes:** ${before.total} / ${MAX_STRIKES}`,
              `**Requested:** +${amount}`,
              `This would exceed the maximum of ${MAX_STRIKES} strikes.`
            ].join("\n")
          );
          return;
        }

        await appendStrikeHistory({
          discordId:
            discordMember.id,
          robloxUsername:
            memberRecord.robloxUsername,
          type:
            strikeType,
          change:
            amount,
          reason,
          staffDiscordId:
            interaction.user.id,
          staff:
            interaction.user.tag ||
            interaction.user.username,
          regiment:
            existingMember.regiment.displayName,
          kompanie:
            existingMember.companyName
        });

        const afterTotal =
          before.total + amount;

        await writeMemberStrikeTotal({
          existingMember,
          total:
            afterTotal
        });

        await interaction.editReply(
          [
            "**Strike Added**",
            "",
            `**Member:** ${discordMember}`,
            `**Type:** ${strikeType}`,
            `**Strikes Added:** ${amount}`,
            `**Reason:** ${reason}`,
            `**Issued By:** ${interaction.user}`,
            "",
            `**Previous Total:** ${before.total} / ${MAX_STRIKES}`,
            `**New Total:** ${afterTotal} / ${MAX_STRIKES}`,
            `**ORBAT:** ${existingMember.regiment.displayName} — ${existingMember.companyName}`,
            getStrikeColumn(
              existingMember.companyName,
              existingMember.regiment.spreadsheetId
            )
              ? `**Strike Column:** F${existingMember.row}`
              : "**Strike Column:** Not displayed on Krümper/Garnison ORBAT"
          ].join("\n")
        );

        return;
      } catch (error) {
        console.error(
          "Failed to add strike:",
          error
        );

        if (
          interaction.deferred ||
          interaction.replied
        ) {
          await interaction.editReply(
            `The strike could not be added: ${error?.message || "Unknown error."}`
          );
        }
        return;
      }
    }

    if (
      interaction.commandName ===
      "removestrike"
    ) {
      try {
        if (
          !interaction.deferred &&
          !interaction.replied
        ) {
          await interaction.deferReply({
            flags:
              MessageFlags.Ephemeral
          });
        }

        console.log(
          "[REMOVE STRIKE] Handler entered"
        );

        const discordMember =
          interaction.options.getUser(
            "discord_member",
            true
          );

        const amount =
          interaction.options.getInteger(
            "amount",
            true
          );

        const reason =
          interaction.options.getString(
            "reason",
            true
          ).trim();

        const existingMember =
          await findMemberByDiscordId(
            discordMember.id
          );

        if (!existingMember) {
          await interaction.editReply(
            "That member was not found in the Grand ORBAT."
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

        const before =
          await applyWeeklyStrikeReduction({
            discordId:
              discordMember.id,
            existingMember,
            memberRecord
          });

        if (
          amount >
          before.total
        ) {
          await interaction.editReply(
            [
              "The strikes were not removed.",
              "",
              `**Current Strikes:** ${before.total} / ${MAX_STRIKES}`,
              `**Requested Removal:** ${amount}`,
              "You cannot remove more strikes than the member currently has."
            ].join("\n")
          );
          return;
        }

        await appendStrikeHistory({
          discordId:
            discordMember.id,
          robloxUsername:
            memberRecord.robloxUsername,
          type:
            "Manual Removal",
          change:
            -amount,
          reason,
          staffDiscordId:
            interaction.user.id,
          staff:
            interaction.user.tag ||
            interaction.user.username,
          regiment:
            existingMember.regiment.displayName,
          kompanie:
            existingMember.companyName
        });

        const afterTotal =
          before.total - amount;

        await writeMemberStrikeTotal({
          existingMember,
          total:
            afterTotal
        });

        await interaction.editReply(
          [
            "**Strikes Removed**",
            "",
            `**Member:** ${discordMember}`,
            `**Strikes Removed:** ${amount}`,
            `**Reason:** ${reason}`,
            `**Removed By:** ${interaction.user}`,
            "",
            `**Previous Total:** ${before.total} / ${MAX_STRIKES}`,
            `**New Total:** ${afterTotal} / ${MAX_STRIKES}`
          ].join("\n")
        );

        return;
      } catch (error) {
        console.error(
          "Failed to remove strike:",
          error
        );

        if (
          interaction.deferred ||
          interaction.replied
        ) {
          await interaction.editReply(
            `The strikes could not be removed: ${error?.message || "Unknown error."}`
          );
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

        const strikeSummary =
          await applyWeeklyStrikeReduction({
            discordId:
              discordMember.id,
            existingMember,
            memberRecord
          });

        const strikeHistoryLines =
          buildStrikeHistoryLines(
            strikeSummary.history,
            10
          );

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
            `**Current Strikes:** ${strikeSummary.total} / ${MAX_STRIKES}`,
            "",
            "**Strike History (latest 10)**",
            ...strikeHistoryLines,
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

    if (interaction.commandName === "organize") {
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

        const inactivityDuration =
          interaction.options
            .getString(
              "inactivity_duration",
              false
            )
            ?.trim() || null;

        const inactivityReason =
          interaction.options
            .getString(
              "inactivity_reason",
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

        const destinationIsGarnison =
          isGarnisonCompany(
            matchedCompany
          );

        const destinationUsesSchuetzenPositions =
          usesSchuetzenPositionStructure(
            newRegiment,
            matchedCompany
          );

        let destinationPosition =
          null;

        let positionWarning =
          null;

        if (
          destinationUsesSchuetzenPositions
        ) {
          if (!newPositionValue) {
            await interaction.editReply(
              [
                "The transfer was not made.",
                "",
                "**The selected Schützen/Jäger company uses a position system.**",
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
              "That destination Schützen/Jäger position is not configured."
            );
            return;
          }
        } else if (newPositionValue) {
          positionWarning =
            destinationIsGarnison
              ? "Position ignored: Garnison Kompanie does not use Company Commander or platoon positions."
              : "Position ignored: the selected destination company does not use the Schützen/Jäger position system.";
        }

        /*
         * Discord slash-command options cannot be conditionally required based
         * on the company dropdown, so these are optional in deploy-commands.js
         * and enforced here only when the destination is Garnison.
         */
        if (destinationIsGarnison) {
          if (!inactivityDuration) {
            await interaction.editReply(
              [
                "The transfer was not made.",
                "",
                "**Garnison Kompanie requires an inactivity duration.**",
                "Enter it in `inactivity_duration` (for example: `2 weeks`, `Until 09/20/2026`, or `Indefinite`)."
              ].join("\n")
            );
            return;
          }

          if (!inactivityReason) {
            await interaction.editReply(
              [
                "The transfer was not made.",
                "",
                "**Garnison Kompanie requires a reason for inactivity.**",
                "Enter it in `inactivity_reason` and try again."
              ].join("\n")
            );
            return;
          }
        }

        const sameRegiment =
          existingMember.regiment.spreadsheetId ===
          newRegiment.spreadsheetId;

        const currentUsesSchuetzenPositions =
          usesSchuetzenPositionStructure(
            existingMember.regiment,
            existingMember.companyName
          );

        const currentPosition =
          currentUsesSchuetzenPositions
            ? getSchuetzenPositionForRow(
                existingMember.row
              )
            : null;

        const sameCompanyName =
          normalizeText(
            existingMember.companyName
          ) === normalizeText(
            matchedCompany
          );

        const sameCompany =
          sameCompanyName &&
          (
            !destinationUsesSchuetzenPositions ||
            (
              currentUsesSchuetzenPositions &&
              currentPosition?.key ===
                destinationPosition?.key
            )
          );

        const internalPositionTransfer =
          sameRegiment &&
          sameCompanyName &&
          destinationUsesSchuetzenPositions &&
          currentUsesSchuetzenPositions &&
          Boolean(currentPosition) &&
          Boolean(destinationPosition) &&
          currentPosition.key !==
            destinationPosition.key;

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

        /*
         * TRANSFER STRIKE PRESERVATION
         *
         * Strike History is authoritative because Krümper and Garnison do
         * not display strikes on their ORBAT sheets. If history exists,
         * carry that total to the destination. If this is an older member
         * with no history yet, fall back to the source ORBAT strike value.
         */
        const transferStrikeHistory =
          await getStrikeHistory(
            discordMember.id
          );

        const transferStrikeTotal =
          transferStrikeHistory.length > 0
            ? calculateStrikeTotal(
                transferStrikeHistory
              )
            : Math.max(
                0,
                Math.min(
                  MAX_STRIKES,
                  Number.parseInt(
                    String(
                      memberRecord.strikes || "0"
                    ),
                    10
                  ) || 0
                )
              );

        memberRecord.strikes =
          transferStrikeTotal;

        console.log(
          "[TRANSFER MEMBER DATA]",
          {
            discordId:
              discordMember.id,
            sourceCompany:
              existingMember.companyName,
            destinationCompany:
              matchedCompany,
            strikes:
              transferStrikeTotal,
            timezone:
              memberRecord.timezone,
            timezoneStorage:
              memberRecord.timezoneStorage
          }
        );

        const previousRank =
          String(memberRecord.rank || "").trim();

        const finalRank =
          requestedNewRank || previousRank;

        try {
          if (
            !usesSchuetzenPositionStructure(
              newRegiment,
              matchedCompany
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
          const updatingGarnisonInactivity =
            destinationIsGarnison &&
            Boolean(
              inactivityDuration &&
              inactivityReason
            );

          if (
            !rankChanged &&
            !updatingGarnisonInactivity
          ) {
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
                destinationIsGarnison
                  ? "Provide both `inactivity_duration` and `inactivity_reason` to update the Garnison inactivity record."
                  : "No spreadsheet cells were changed."
              ].join("\n")
            );
            return;
          }

          if (rankChanged) {
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
          }

          if (updatingGarnisonInactivity) {
            await writeGarnisonInactivity({
              spreadsheetId:
                existingMember.regiment.spreadsheetId,
              sheetName:
                existingMember.companyName,
              row:
                existingMember.row,
              duration:
                inactivityDuration,
              reason:
                inactivityReason
            });
          }

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
            destinationIsGarnison && !rankChanged
              ? "Garnison inactivity information updated successfully through /transfer."
              : "Member transfer information updated successfully.",
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

          if (destinationIsGarnison) {
            rankOnlyReply.push(
              `**Duration of Inactivity:** ${inactivityDuration}`,
              `**Reason for Inactivity:** ${inactivityReason}`
            );
          }

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
         * Attendance is company-specific except for an internal
         * Jäger/Schützen position change within the SAME Kompanie.
         */
        const attendanceToPreserve =
          internalPositionTransfer
            ? await getStandardAttendanceRecord({
                spreadsheetId:
                  existingMember.regiment.spreadsheetId,
                sheetName:
                  existingMember.companyName,
                row:
                  existingMember.row
              })
            : null;

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
            strikes:
              memberRecord.strikes,
            position:
              destinationPosition
          });

        if (destinationIsGarnison) {
          await writeGarnisonInactivity({
            spreadsheetId:
              newRegiment.spreadsheetId,
            sheetName:
              matchedCompany,
            row:
              destinationRow,
            duration:
              inactivityDuration,
            reason:
              inactivityReason
          });
        }

        if (
          internalPositionTransfer &&
          attendanceToPreserve
        ) {
          await writeStandardAttendanceRecord({
            spreadsheetId:
              newRegiment.spreadsheetId,
            sheetName:
              matchedCompany,
            row:
              destinationRow,
            attendance:
              attendanceToPreserve
          });
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

        if (
          memberRecord.timezone ||
          memberRecord.timezoneStorage
        ) {
          try {
            /*
             * Prefer the hidden canonical timezone when available.
             * This avoids reparsing an abbreviation such as PST/PDT/EST.
             */
            const storedTimezone =
              String(
                memberRecord.timezoneStorage || ""
              ).trim();

            const timezoneForProcessing =
              storedTimezone.startsWith("IANA:")
                ? storedTimezone.substring(5)
                : storedTimezone.startsWith("OFFSET:")
                  ? storedTimezone.substring(7)
                  : (
                      storedTimezone ||
                      memberRecord.timezone
                    );

            const timezoneResult =
              await processTimezoneWithAppsScript({
                spreadsheetId:
                  newRegiment.spreadsheetId,
                sheetName:
                  matchedCompany,
                row:
                  destinationRow,
                timezone:
                  timezoneForProcessing,
                strikes:
                  memberRecord.strikes
              });

            /*
             * Enforce the DESTINATION layout after Apps Script:
             *
             * Krümper/Garnison:
             *   F = timezone, G = storage, no displayed strikes
             *
             * All other companies:
             *   F = strikes, G = timezone, H = storage
             */
            await enforceOrbatTimezoneLayout({
              spreadsheetId:
                newRegiment.spreadsheetId,
              sheetName:
                matchedCompany,
              row:
                destinationRow,
              timezone:
                timezoneResult?.displayValue ||
                memberRecord.timezone,
              ianaTimezone:
                timezoneResult?.storageValue ||
                memberRecord.timezoneStorage ||
                ""
            });

            console.log(
              "[TRANSFER DESTINATION LAYOUT ENFORCED]",
              {
                discordId:
                  discordMember.id,
                company:
                  matchedCompany,
                row:
                  destinationRow,
                strikes:
                  memberRecord.strikes,
                timezone:
                  timezoneResult?.displayValue ||
                  memberRecord.timezone,
                storage:
                  timezoneResult?.storageValue ||
                  memberRecord.timezoneStorage ||
                  ""
              }
            );
          } catch (webhookError) {
            console.error(
              "Transfer completed, but timezone processing failed:"
            );
            console.error(webhookError);

            timezoneWarning =
              webhookError?.message ||
              "The Apps Script webhook failed.";

            /*
             * Even if Apps Script is unavailable, keep the destination
             * sheet structurally correct using the values we already read
             * from the source ORBAT.
             */
            try {
              await enforceOrbatTimezoneLayout({
                spreadsheetId:
                  newRegiment.spreadsheetId,
                sheetName:
                  matchedCompany,
                row:
                  destinationRow,
                timezone:
                  memberRecord.timezone,
                ianaTimezone:
                  memberRecord.timezoneStorage ||
                  ""
              });
            } catch (layoutError) {
              console.error(
                "Transfer destination layout fallback failed:"
              );
              console.error(layoutError);
            }
          }
        } else {
          /*
           * There is no timezone to process, but strike-enabled destination
           * sheets must still receive the preserved strike total in F.
           */
          const destinationStrikeColumn =
            getStrikeColumn(
              matchedCompany,
              newRegiment.spreadsheetId
            );

          if (destinationStrikeColumn) {
            const safeDestinationSheet =
              escapeSheetName(
                matchedCompany
              );

            await sheets.spreadsheets.values.update({
              spreadsheetId:
                newRegiment.spreadsheetId,
              range:
                `${safeDestinationSheet}!${destinationStrikeColumn}${destinationRow}`,
              valueInputOption:
                "RAW",
              requestBody: {
                values: [[
                  memberRecord.strikes
                ]]
              }
            });
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

        await syncMasterRosterMember({
          discordId:
            discordMember.id,
          robloxUsername:
            memberRecord.robloxUsername,
          rank:
            finalRank,
          regiment:
            newRegiment.displayName,
          kompanie:
            matchedCompany,
          position:
            getMasterRosterPosition({
              regiment:
                newRegiment,
              company:
                matchedCompany,
              row:
                destinationRow,
              explicitPosition:
                destinationPosition
            }),
          active: "Yes"
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
          `**Strikes Carried:** ${memberRecord.strikes} / ${MAX_STRIKES}`,
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

        if (positionWarning) {
          replyLines.push(
            "",
            positionWarning
          );
        }

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
                memberRecord.timezone,
              strikes:
                memberRecord.strikes
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
                memberRecord.timezone,
              strikes:
                memberRecord.strikes
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

          await syncMasterRosterMember({
            discordId:
              discordMember.id,
            robloxUsername:
              memberRecord.robloxUsername,
            rank:
              newRank,
            regiment:
              existingMember.regiment.displayName,
            kompanie:
              destinationCompany,
            position:
              getMasterRosterPosition({
                regiment:
                  existingMember.regiment,
                company:
                  destinationCompany,
                row:
                  destinationRow
              }),
            active: "Yes"
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

        await syncMasterRosterMember({
          discordId:
            discordMember.id,
          robloxUsername:
            memberRecord.robloxUsername,
          rank:
            newRank,
          regiment:
            existingMember.regiment.displayName,
          kompanie:
            existingMember.companyName,
          position:
            getMasterRosterPosition({
              regiment:
                existingMember.regiment,
              company:
                existingMember.companyName,
              row:
                sortedRow
            }),
          active: "Yes"
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

        /*
         * Keep the member's permanent Army history, but mark them inactive.
         */
        await markMasterRosterMemberInactive({
          discordId:
            discordMember.id,
          robloxUsername:
            memberRecord.robloxUsername,
          rank:
            memberRecord.rank
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
            `Cleared the member data from spreadsheet row ${existingMember.row}.`
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

      const companyUsesSchuetzenPositions =
        usesSchuetzenPositionStructure(
          regiment,
          matchedCompany
        );

      let schuetzenPosition =
        null;

      if (
        companyUsesSchuetzenPositions
      ) {
        if (!positionValue) {
          await interaction.editReply(
            [
              "The member was not added.",
              "",
              "**The selected Schützen/Jäger company uses a position system.**",
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
            "That Schützen/Jäger position is not configured."
          );
          return;
        }
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
          !usesSchuetzenPositionStructure(
            regiment,
            matchedCompany
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

      /*
       * Final safeguard: re-apply the destination sheet's column layout
       * after the Apps Script webhook. This makes /addmember authoritative
       * even if Railway is still pointing at an older Apps Script deployment.
       */
      await enforceOrbatTimezoneLayout({
        spreadsheetId: regiment.spreadsheetId,
        sheetName: matchedCompany,
        row,
        timezone:
          timezoneResult?.displayValue ||
          timezone,
        ianaTimezone:
          timezoneResult?.storageValue ||
          ""
      });

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

      /*
       * Synchronize the army-wide Master Roster only after the ORBAT
       * member has been created and their final sorted row is known.
       */
      await syncMasterRosterMember({
        discordId:
          discordMember.id,
        robloxUsername,
        rank,
        regiment:
          regiment.displayName,
        kompanie:
          matchedCompany,
        position:
          getMasterRosterPosition({
            regiment,
            company:
              matchedCompany,
            row,
            explicitPosition:
              schuetzenPosition
          }),
        active: "Yes"
      });

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
        `Written to spreadsheet row ${row}. Timezone column: ${getTimezoneColumnLayout(matchedCompany, regiment.spreadsheetId).timezoneColumn}${row}.`
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
          `The original timezone remains in ${getTimezoneColumnLayout(matchedCompany, regiment.spreadsheetId).timezoneColumn}${row}.`
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