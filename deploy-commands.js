import "dotenv/config";

import {
  REST,
  Routes,
  SlashCommandBuilder
} from "discord.js";

function cleanEnvironmentValue(value) {
  return String(value || "")
    .replace(/[\u2028\u2029\uFEFF]/g, "")
    .trim();
}

const DISCORD_TOKEN = cleanEnvironmentValue(
  process.env.DISCORD_TOKEN
);

const CLIENT_ID = cleanEnvironmentValue(
  process.env.CLIENT_ID
);

const GUILD_ID = cleanEnvironmentValue(
  process.env.GUILD_ID
);

if (!DISCORD_TOKEN) {
  throw new Error(
    "DISCORD_TOKEN is missing from the .env file."
  );
}

if (!CLIENT_ID) {
  throw new Error(
    "CLIENT_ID is missing from the .env file."
  );
}

if (!GUILD_ID) {
  throw new Error(
    "GUILD_ID is missing from the .env file."
  );
}

const REGIMENT_CHOICES = [
  {
    name: "11. Erstes Schlesisches Infanterie-Regiment",
    value: "11_schlesisches"
  },
  {
    name:
      "6. Erstes Westpreußisches Infanterie-Regiment",
    value: "6_westpreussisches"
  },
  {
    name: "Ostpreußisches Jäger-Bataillon",
    value: "ostpreussisches_jaeger"
  }
];

const RANK_CHOICES = [
  { name: "Rekrut", value: "Rekrut" },
  { name: "Soldat", value: "Soldat" },
  { name: "Obersoldat", value: "Obersoldat" },
  { name: "Gefreiter", value: "Gefreiter" },
  { name: "Obergefreiter", value: "Obergefreiter" },
  { name: "Vizekorporal", value: "Vizekorporal" },
  { name: "Korporal", value: "Korporal" },
  { name: "Sergeant", value: "Sergeant" },
  { name: "Feldwebel", value: "Feldwebel" },
  { name: "Fähnrich", value: "Fähnrich" },
  {
    name: "Sekonde-Lieutenant",
    value: "Sekonde-Lieutenant"
  },
  {
    name: "Premier-Lieutenant",
    value: "Premier-Lieutenant"
  },
  { name: "Kapitän", value: "Kapitän" },
  {
    name: "Stabs-Kapitän",
    value: "Stabs-Kapitän"
  },
  { name: "Major", value: "Major" },
  {
    name: "Oberst-Lieutenant",
    value: "Oberst-Lieutenant"
  },
  { name: "Oberst", value: "Oberst" }
];

/*
 * IMPORTANT:
 * These must match the exact text used by the Google Sheets attendance
 * dropdown. If your sheet uses different wording, change only this list
 * and the ATTENDANCE_STATUS_CHOICES list in index.js.
 */
const ATTENDANCE_CHOICES = [
  { name: "DM", value: "DM" },
  { name: "RSVP", value: "RSVP" },
  { name: "MAYB", value: "MAYB" },
  { name: "NO", value: "NO" },
  { name: "PRES", value: "PRES" },
  { name: "EXC", value: "EXC" },
  { name: "AWOL", value: "AWOL" },
  { name: "LEFT", value: "LEFT" }
];

const commands = [
  new SlashCommandBuilder()
    .setName("addmember")
    .setDescription(
      "Adds a RoVer-verified member to a regiment ORBAT spreadsheet."
    )
    .addUserOption(option =>
      option
        .setName("discord_member")
        .setDescription(
          "Select the member's verified Discord account."
        )
        .setRequired(true)
    )
    .addStringOption(option =>
      option
        .setName("regiment")
        .setDescription(
          "Select the member's regiment."
        )
        .setRequired(true)
        .addChoices(...REGIMENT_CHOICES)
    )
    .addStringOption(option =>
      option
        .setName("company")
        .setDescription(
          "Select a company from the chosen regiment spreadsheet."
        )
        .setRequired(true)
        .setAutocomplete(true)
    )
    .addStringOption(option =>
      option
        .setName("rank")
        .setDescription(
          "Select the member's rank."
        )
        .setRequired(true)
        .addChoices(...RANK_CHOICES)
    )
    .addStringOption(option =>
      option
        .setName("timezone")
        .setDescription(
          "Enter a city, state, country, timezone, or UTC offset."
        )
        .setRequired(true)
        .setMinLength(1)
        .setMaxLength(100)
    )
    .toJSON(),

  new SlashCommandBuilder()
    .setName("transfer")
    .setDescription(
      "Transfers a Grand ORBAT member to another regiment or company."
    )
    .addUserOption(option =>
      option
        .setName("discord_member")
        .setDescription(
          "Select the Discord member to transfer."
        )
        .setRequired(true)
    )
    .addStringOption(option =>
      option
        .setName("new_regiment")
        .setDescription(
          "Select the member's destination regiment."
        )
        .setRequired(true)
        .addChoices(...REGIMENT_CHOICES)
    )
    .addStringOption(option =>
      option
        .setName("new_company")
        .setDescription(
          "Select the member's destination company."
        )
        .setRequired(true)
        .setAutocomplete(true)
    )
    .addStringOption(option =>
      option
        .setName("new_rank")
        .setDescription(
          "Optionally select the member's new rank."
        )
        .setRequired(false)
        .addChoices(...RANK_CHOICES)
    )
    .toJSON(),

  new SlashCommandBuilder()
    .setName("rank")
    .setDescription(
      "Changes a Grand ORBAT member's rank."
    )
    .addUserOption(option =>
      option
        .setName("discord_member")
        .setDescription(
          "Select the Discord member whose rank should be changed."
        )
        .setRequired(true)
    )
    .addStringOption(option =>
      option
        .setName("new_rank")
        .setDescription(
          "Select the member's new rank."
        )
        .setRequired(true)
        .addChoices(...RANK_CHOICES)
    )
    .toJSON(),

  new SlashCommandBuilder()
    .setName("removemember")
    .setDescription(
      "Removes a member from the Grand ORBAT."
    )
    .addUserOption(option =>
      option
        .setName("discord_member")
        .setDescription(
          "Select the Discord member to remove."
        )
        .setRequired(true)
    )
    .toJSON(),

  new SlashCommandBuilder()
    .setName("orginize")
    .setDescription(
      "Organizes a company roster from highest rank to lowest rank."
    )
    .addStringOption(option =>
      option
        .setName("regiment")
        .setDescription(
          "Select the regiment."
        )
        .setRequired(true)
        .addChoices(...REGIMENT_CHOICES)
    )
    .addStringOption(option =>
      option
        .setName("company")
        .setDescription(
          "Select the company to organize."
        )
        .setRequired(true)
        .setAutocomplete(true)
    )
    .toJSON(),

  new SlashCommandBuilder()
    .setName("attendance")
    .setDescription(
      "Starts a multi-member company attendance entry."
    )
    .addStringOption(option =>
      option
        .setName("regiment")
        .setDescription(
          "Select the regiment."
        )
        .setRequired(true)
        .addChoices(...REGIMENT_CHOICES)
    )
    .addStringOption(option =>
      option
        .setName("company")
        .setDescription(
          "Select the company."
        )
        .setRequired(true)
        .setAutocomplete(true)
    )
    .toJSON()
];

const rest = new REST({
  version: "10"
}).setToken(DISCORD_TOKEN);

async function deployCommands() {
  try {
    console.log(
      "Registering /addmember, /transfer, /rank, /removemember, /attendance, and /orginize..."
    );

    await rest.put(
      Routes.applicationGuildCommands(
        CLIENT_ID,
        GUILD_ID
      ),
      {
        body: commands
      }
    );

    console.log(
      "Commands registered successfully."
    );
  } catch (error) {
    console.error(
      "Failed to register Discord commands:"
    );
    console.error(error);
    process.exitCode = 1;
  }
}

deployCommands();