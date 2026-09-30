const fs = require("fs");
const { PublicClientApplication } = require("@azure/msal-node");

const config = JSON.parse(fs.readFileSync("./config.json", "utf8"));

const pca = new PublicClientApplication({
  auth: {
    clientId: config.clientId,
    authority: `https://login.microsoftonline.com/${config.tenantId}`,
  },
});

async function main() {
  console.log("Starting Microsoft login...");

  const result = await pca.acquireTokenByDeviceCode({
    scopes: ["Chat.ReadBasic"],
    deviceCodeCallback: (response) => {
      console.log("\n========== RAW DEVICE RESPONSE ==========");
      console.log(JSON.stringify(response, null, 2));
      console.log("=========================================\n");
    },
  });

  console.log("Login successful!");
  console.log("Signed in as:", result.account?.username);

  const response = await fetch("https://graph.microsoft.com/v1.0/me/chats", {
    headers: {
      Authorization: `Bearer ${result.accessToken}`,
    },
  });

  const data = await response.json();

  if (!response.ok) {
    console.error("Graph API error:");
    console.error(JSON.stringify(data, null, 2));
    return;
  }

  for (const chat of data.value) {
    console.log("\n-----------------------------");
    console.log("Chat name:", chat.topic || "(no topic)");
    console.log("Chat type:", chat.chatType);
    console.log("Chat ID:", chat.id);
  }
}

main().catch((error) => {
  console.error("\nAUTH ERROR:");
  console.error(error);
});
