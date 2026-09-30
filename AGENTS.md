# Developer & Agent Guide: teams-morning

This document outlines project conventions, operational workflows, and security guardrails for working on the `teams-morning` codebase.

---

## Project Overview

A lightweight Node.js automation utility using `@azure/msal-node` and Microsoft Graph API to send daily morning notifications to a Microsoft Teams chat.

### Core Architectural Concepts
- **Authentication**: Uses MSAL PublicClientApplication device-code flow (`pca.acquireTokenByDeviceCode`) and silent acquisition (`pca.acquireTokenSilent`).
- **Token Cache Persistence**: Tokens are serialized/deserialized to `token-cache.json` using MSAL's `cachePlugin`.
- **Deduplication / State**: `last-sent.txt` stores the date formatted as `YYYY-MM-DD` of the last successful transmission.
- **Weekend Gate**: `isWeekend()` enforces that execution terminates gracefully without sending on Saturdays and Sundays.
- **Time Gate**: `isBefore1030()` enforces that execution terminates gracefully without sending if initiated at or after 10:30 AM local time.

---

## Commands & Scripts

All commands should be executed from the project root (`c:\teams-morning`):

- **Start / Send Morning Message**:
  ```bash
  npm start
  # or
  npm run send
  ```
- **List Teams Chats & Retrieve Chat IDs**:
  ```bash
  npm run get-chats
  ```
- **Test Device Code Flow Directly**:
  ```bash
  npm run test-device-code
  ```

---

## Configuration & Secrets Handling

- **Required Config**: `config.json` must be present at runtime with:
  - `clientId`: Azure Active Directory App ID
  - `tenantId`: Azure Active Directory Tenant ID
  - `chatId`: Target Microsoft Teams chat thread ID
- **Safety Policy**:
  - NEVER commit or hardcode actual values of `clientId`, `tenantId`, `chatId`, or the contents of `token-cache.json` into git or documentation.
  - Maintain `config.example.json` with dummy values for reference.
  - The `.gitignore` must ignore `config.json`, `token-cache.json`, `last-sent.txt`, `node_modules/`, and logs.

---

## Required Microsoft Graph Scopes

When modifying scopes in [send-teams.js](file:///c:/teams-morning/send-teams.js) or [get-chats.js](file:///c:/teams-morning/get-chats.js), ensure they match the delegated permissions configured in the Azure AD application registration:
- `User.Read`
- `Chat.ReadBasic`
- `ChatMessage.Send`
