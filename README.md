# Teams Morning Message Bot 🌅

An automated script that sends a daily morning greeting to a specified Microsoft Teams chat using Microsoft Graph API and MSAL Node (`@azure/msal-node`).

---

## Features

- **OAuth 2.0 Device Code Flow**: Authenticates directly with your Microsoft 365 work/school account without exposing personal passwords.
- **Persistent Token Cache**: Caches tokens locally in `token-cache.json` using MSAL's cache plugin so you only need to complete the interactive device code login once. Silent token refreshes occur automatically thereafter.
- **Time Window Restriction**: Ensures messages are only sent before **10:30 AM** (configurable in [send-teams.js](file:///c:/teams-morning/send-teams.js)).
- **Once-Per-Day Guard**: Stores the last sent date in [last-sent.txt](file:///c:/teams-morning/last-sent.txt) to prevent duplicate messages if triggered multiple times.
- **Chat Discovery Tool**: Includes a helper script ([get-chats.js](file:///c:/teams-morning/get-chats.js)) to easily discover and copy your Teams chat IDs.

---

## Prerequisites

1. **Node.js**: v18.0.0 or higher recommended.
2. **Microsoft 365 Account**: An account with access to Microsoft Teams.
3. **Azure AD App Registration**:
   - Register an application in the [Azure Portal (Entra ID)](https://portal.azure.com/).
   - Under **Authentication**:
     - Platform configurations: Add a platform -> **Mobile and desktop applications** (e.g., redirect URI `https://login.microsoftonline.com/common/oauth2/nativeclient`).
     - Advanced settings -> **Allow public client flows**: Set **Enable the following mobile and desktop flows** to **Yes**.
   - Under **API Permissions**, add the following **Delegated** Microsoft Graph permissions:
     - `User.Read`
     - `Chat.ReadBasic`
     - `ChatMessage.Send`
   - Grant admin consent if required by your organization.

---

## Installation & Setup

1. **Install Dependencies**:
   ```bash
   npm install
   ```

2. **Configure Credentials**:
   Copy [config.example.json](file:///c:/teams-morning/config.example.json) to `config.json`:
   ```bash
   cp config.example.json config.json
   ```

3. **Fill in [config.json](file:///c:/teams-morning/config.json)**:
   ```json
   {
     "clientId": "your-azure-app-client-id",
     "tenantId": "your-azure-ad-tenant-id",
     "chatId": "19:your-chat-thread-id@thread.v2"
   }
   ```

---

## Discovering Your Chat ID

If you do not know the target `chatId`:

1. Run the chat discovery script:
   ```bash
   npm run get-chats
   ```
2. Follow the prompt to visit `https://microsoft.com/devicelogin` and enter the displayed code.
3. Once logged in, the script will output all chats associated with your account:
   ```text
   -----------------------------
   Chat name: Morning Standup
   Chat type: group
   Chat ID: 19:4904ba34d01442bab28b7dafb054386e@thread.v2
   ```
4. Copy the relevant `Chat ID` into `chatId` inside `config.json`.

---

## Running the Bot

### Send Morning Message
```bash
npm start
# or
npm run send
```

**Execution Flow**:
1. Checks that `clientId`, `tenantId`, and `chatId` exist in `config.json`.
2. Checks that current time is before **10:30 AM**.
3. Checks [last-sent.txt](file:///c:/teams-morning/last-sent.txt) to confirm no message was already sent today (`YYYY-MM-DD`).
4. Attempts silent token acquisition from [token-cache.json](file:///c:/teams-morning/token-cache.json). If no valid session exists, initiates device code login in the terminal.
5. Sends the greeting via Microsoft Graph API (`POST /v1.0/chats/{chatId}/messages`).
6. Updates `last-sent.txt` with today's date upon successful delivery.

### Diagnostic / Test Scripts
- **Test Device Code Flow**:
  ```bash
  npm run test-device-code
  ```
  Directly verifies connectivity and response from the Microsoft OAuth device code endpoint.

---

## Scheduling Daily Automation

### Windows Task Scheduler
To automatically trigger the message every weekday morning:

1. Open **Task Scheduler** (`taskschd.msc`).
2. Click **Create Task...**.
3. **General**:
   - Name: `Teams Morning Greeting`
   - Select *Run only when user is logged on* (or *Run whether user is logged on or not* with stored credentials).
4. **Triggers**:
   - New -> Daily or Weekly (e.g., Monday through Friday at 9:00 AM).
5. **Actions**:
   - New -> Action: *Start a program*.
   - Program/script: `node` (or full path like `C:\Program Files\nodejs\node.exe`).
   - Add arguments: `send-teams.js`
   - Start in: `C:\teams-morning`
6. Save and test by right-clicking the task and selecting **Run**.

---

## Project Structure

```
teams-morning/
├── config.json             # Local credentials (gitignored)
├── config.example.json     # Configuration template
├── get-chats.js            # Helper script to list Teams chats & IDs
├── last-sent.txt           # Tracks last successful send date (gitignored)
├── package.json            # Node.js dependencies & scripts
├── send-teams.js           # Core execution script
├── test-device-code.js     # Low-level OAuth endpoint test script
├── token-cache.json        # Serialized MSAL token cache (gitignored)
└── README.md               # Project documentation
```

---

## Security Notes

- **Never commit credentials**: `config.json` and `token-cache.json` contain sensitive tenant IDs, client IDs, and active session tokens. Keep them protected in [.gitignore](file:///c:/teams-morning/.gitignore).
- **Public Client Permissions**: The Azure AD app registration uses public client device code authentication; never embed client secrets in client-side or distributed scripts.
