const fs = require("fs");

const config = JSON.parse(fs.readFileSync("./config.json", "utf8"));

console.log(config);

async function main() {
  const body = new URLSearchParams({
    client_id: config.clientId,
    scope: "User.Read offline_access openid profile",
  });

  const response = await fetch(
    `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/devicecode`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body,
    },
  );

  const text = await response.text();

  console.log("HTTP STATUS:", response.status);
  console.log("RESPONSE:");
  console.log(text);
}

main().catch(console.error);
