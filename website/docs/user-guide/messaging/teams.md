---
sidebar_position: 5
title: "Microsoft Teams"
description: "Set up Hermes Agent as a Microsoft Teams bot using the Teams CLI"
---

# Microsoft Teams Setup

Connect Hermes Agent to Microsoft Teams as a bot. Teams requires a public HTTPS endpoint to deliver messages, so you'll need either a dev tunnel (local dev) or a public server (production).

## Overview

| Component | Value |
|-----------|-------|
| **Library** | `microsoft-teams-apps` |
| **Connection** | Webhook — public HTTPS endpoint required |
| **Auth required** | Azure AD App ID + Client Secret + Tenant ID |
| **Webhook port** | 3978 (default) |
| **User identification** | AAD object IDs (find yours with `teams status --verbose`) |

---

## Step 1: Install the Teams CLI

The `@microsoft/teams.cli` automates bot registration — no Azure portal needed.

```bash
npm install -g @microsoft/teams.cli@preview
teams login
```

To verify your login and find your own AAD object ID (useful for `TEAMS_ALLOWED_USERS`):

```bash
teams status --verbose
```

---

## Step 2: Expose Port 3978

Teams cannot reach `localhost`, so you need a tunnel. Install [devtunnel](https://learn.microsoft.com/en-us/azure/developer/dev-tunnels/get-started):

```bash
devtunnel create hermes-bot --allow-anonymous
devtunnel port create hermes-bot -p 3978 --protocol auto
devtunnel host hermes-bot
```

Copy the `https://` URL from the output — you'll use it in the next step.

---

## Step 3: Create the Bot

```bash
teams app create \
  --name "Hermes" \
  --endpoint "https://<your-tunnel-url>/api/messages"
```

The CLI outputs your `CLIENT_ID`, `CLIENT_SECRET`, and `TENANT_ID`. Save them.

---

## Step 4: Configure Environment Variables

Add to `~/.hermes/.env`:

```bash
TEAMS_CLIENT_ID=<your-client-id>
TEAMS_CLIENT_SECRET=<your-client-secret>
TEAMS_TENANT_ID=<your-tenant-id>

# Restrict access to your AAD object ID (from `teams status --verbose`)
TEAMS_ALLOWED_USERS=<your-aad-object-id>
```

---

## Step 5: Run with Docker

```bash
HERMES_UID=$(id -u) HERMES_GID=$(id -g) docker compose up -d gateway
```

The gateway listens on port 3978. With `network_mode: host`, no extra port mapping is needed.

Check logs:

```bash
docker logs -f hermes
```

---

## Step 6: Install the App in Teams

```bash
teams app install --id <teamsAppId>
```

Then DM the bot in Teams — it's ready.

---

## Configuration Reference

| Variable | Description |
|----------|-------------|
| `TEAMS_CLIENT_ID` | Azure AD App (client) ID |
| `TEAMS_CLIENT_SECRET` | Azure AD client secret |
| `TEAMS_TENANT_ID` | Azure AD tenant ID |
| `TEAMS_ALLOWED_USERS` | Comma-separated AAD object IDs |
| `TEAMS_ALLOW_ALL_USERS` | Set `true` to skip the allowlist |
| `TEAMS_HOME_CHANNEL` | Default channel/chat ID for cron delivery |
| `TEAMS_HOME_CHANNEL_NAME` | Display name for the home channel |
| `TEAMS_PORT` | Webhook port (default: `3978`) |

Alternatively, configure via `~/.hermes/config.yaml`:

```yaml
platforms:
  teams:
    enabled: true
    extra:
      client_id: "your-client-id"
      client_secret: "your-secret"
      tenant_id: "your-tenant-id"
      port: 3978
```

---

## Production Deployment

For a permanent server, skip devtunnel and point the bot's messaging endpoint directly at your host:

```
https://your-domain.com/api/messages
```

Update the endpoint in the Teams CLI:

```bash
teams app update --id <teamsAppId> --endpoint "https://your-domain.com/api/messages"
```
