const fs = require("fs");
const {
  PublicClientApplication,
  InteractionRequiredAuthError,
} = require("@azure/msal-node");

const config = JSON.parse(fs.readFileSync("./config.json", "utf8"));

const CACHE_FILE = "./token-cache.json";
const LAST_SENT_FILE = "./last-sent.txt";

const GRAPH_SCOPES = ["User.Read", "Chat.ReadBasic", "ChatMessage.Send"];

const cachePlugin = {
  beforeCacheAccess: async (cacheContext) => {
    if (fs.existsSync(CACHE_FILE)) {
      cacheContext.tokenCache.deserialize(fs.readFileSync(CACHE_FILE, "utf8"));
    }
  },

  afterCacheAccess: async (cacheContext) => {
    if (cacheContext.cacheHasChanged) {
      fs.writeFileSync(CACHE_FILE, cacheContext.tokenCache.serialize());
    }
  },
};

const pca = new PublicClientApplication({
  auth: {
    clientId: config.clientId,
    authority: `https://login.microsoftonline.com/${config.tenantId}`,
  },
  cache: {
    cachePlugin,
  },
});

function getToday() {
  const now = new Date();

  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(
    2,
    "0",
  )}-${String(now.getDate()).padStart(2, "0")}`;
}

function isBefore1030() {
  const now = new Date();

  const minutes = now.getHours() * 60 + now.getMinutes();

  return minutes < 10 * 60 + 30;
}

function isWeekend(now = new Date()) {
  const day = now.getDay();
  return day === 0 || day === 6; // 0 = Sunday, 6 = Saturday
}

async function getToken() {
  const accounts = await pca.getTokenCache().getAllAccounts();

  // Try existing login first
  if (accounts.length > 0) {
    try {
      console.log("Using existing Microsoft login...");

      return await pca.acquireTokenSilent({
        account: accounts[0],
        scopes: GRAPH_SCOPES,
      });
    } catch (error) {
      if (!(error instanceof InteractionRequiredAuthError)) {
        throw error;
      }

      console.log("Existing login requires authentication again.");
    }
  }

  // First-time authentication
  console.log("Starting Microsoft login...");

  return await pca.acquireTokenByDeviceCode({
    scopes: GRAPH_SCOPES,

    deviceCodeCallback: (response) => {
      console.log("\n=================================");
      console.log(response.message);
      console.log("=================================\n");
    },
  });
}

async function main() {
  // -------------------------------
  // Validate configuration
  // -------------------------------

  if (!config.clientId) {
    throw new Error("clientId is missing in config.json");
  }

  if (!config.tenantId) {
    throw new Error("tenantId is missing in config.json");
  }

  if (!config.chatId) {
    throw new Error("chatId is missing in config.json");
  }

  // -------------------------------
  // Weekend check
  // -------------------------------

  if (isWeekend()) {
    console.log("Today is a weekend (Saturday or Sunday). Skipping message.");
    return;
  }

  // -------------------------------
  // Time check
  // -------------------------------

  if (!isBefore1030()) {
    console.log("It is already 10:30 AM or later. Skipping message.");
    return;
  }

  // -------------------------------
  // Once-per-day check
  // -------------------------------

  const today = getToday();

  if (fs.existsSync(LAST_SENT_FILE)) {
    const lastSent = fs.readFileSync(LAST_SENT_FILE, "utf8").trim();

    if (lastSent === today) {
      console.log("Today's message was already sent. Skipping.");
      return;
    }
  }

  // -------------------------------
  // Get Microsoft token
  // -------------------------------

  const token = await getToken();

  // -------------------------------
  // Send Teams message
  // -------------------------------

  console.log("Sending Teams message...");

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(
      config.chatId,
    )}/messages`,
    {
      method: "POST",

      headers: {
        Authorization: `Bearer ${token.accessToken}`,
        "Content-Type": "application/json",
      },

      body: JSON.stringify({
        body: {
          content: "🌅 Good morning! Have a productive day 🚀",
        },
      }),
    },
  );

  const data = await response.json();

  if (!response.ok) {
    console.error("\nTeams API error:");
    console.error(JSON.stringify(data, null, 2));
    process.exit(1);
  }

  // -------------------------------
  // Mark as sent ONLY after success
  // -------------------------------

  fs.writeFileSync(LAST_SENT_FILE, today);

  console.log("✅ Teams message sent successfully.");
  console.log(`✅ Recorded ${today} in last-sent.txt`);
}

main().catch((error) => {
  console.error("\n❌ Error:");
  console.error(error);
  process.exit(1);
});
