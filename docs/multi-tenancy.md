# Multi-Tenancy Guide

This guide covers how to host the MS 365 MCP Server as a shared service for multiple customers across different Azure AD / Entra ID tenants.

## Architecture Overview

The server's HTTP mode (`--http`) already supports multi-tenant, multi-user access. Each incoming request carries the customer's own Microsoft access token, and the server uses per-request token isolation via `AsyncLocalStorage` to ensure one customer's token is never used for another's request.

**Request flow:**

1. Your platform authenticates each customer against Microsoft (OAuth dance) and stores their access + refresh tokens.
2. Each MCP request includes the customer's token as `Authorization: Bearer <access_token>`.
3. The server's bearer token middleware extracts the token and stores it in the request context.
4. `GraphClient` checks the per-request context first when making Graph API calls — the shared `AuthManager` singleton is never consulted.
5. If the access token expires mid-request (401 from Microsoft), the server uses the refresh token from the request context to obtain a new one.

No code changes are needed for multi-tenancy. It works today.

## App Registration Setup

You need **one** Azure App Registration that users from multiple tenants can authenticate against.

### Tenant ID Configuration

The `MS365_MCP_TENANT_ID` (or `--cloud`) setting controls the OAuth authority URL and determines who can sign in:

| Value | Who can sign in | Use case |
|---|---|---|
| A specific GUID (e.g., `a1b2c3d4-...`) | Only users from that one Entra ID tenant | Single-company internal deployment |
| `organizations` | Users from any Entra ID tenant (work/school accounts) | Multi-tenant SaaS, business customers only |
| `common` (default) | Any Entra ID tenant + personal Microsoft accounts (Outlook.com, Hotmail) | Multi-tenant SaaS including solopreneurs/freelancers |
| `consumers` | Personal Microsoft accounts only | Consumer-facing apps |

For a SaaS product serving businesses, `organizations` is the most appropriate. The default `common` is more permissive and also works — it just additionally allows personal Microsoft accounts.

### Making an Existing App Registration Multi-Tenant

If you have a single-tenant App Registration:

1. Azure Portal > your App Registration > **Authentication** > **Supported account types**
2. Change to "Accounts in any organizational directory" (`organizations`) or add personal accounts (`common`)
3. Save

This does **not** invalidate existing OAuth connections. Existing tokens and refresh tokens continue to work. The change only allows new tenants to authenticate.

## Admin Consent

When a user from a new tenant signs in for the first time, their organization must consent to the permissions your app requests.

### How Consent Works

- **User-consentable permissions** (e.g., `Mail.Read`, `Calendars.Read`, `User.Read`): The user sees a consent prompt on first login and can approve it themselves. No admin needed.
- **Admin-consent-required permissions** (e.g., `Mail.ReadWrite`, Teams/SharePoint scopes): A tenant admin must grant consent for the entire organization.

You can check which permissions require admin consent in your App Registration under **API permissions** — there's an "Admin consent required" column.

### Admin Consent URL

Provide customers with an admin consent link as part of your onboarding flow:

```
https://login.microsoftonline.com/common/adminconsent?client_id=YOUR_CLIENT_ID
```

An admin clicks it, signs in, reviews the permissions, and approves for their entire organization. This is a one-time step per tenant.

Alternatively, admins can grant consent through the Azure Portal: Entra ID > Enterprise Applications > find your app > grant consent.

## Managing Permission Scope with Presets

Azure admin consent is all-or-nothing for the permissions your app requests. If your App Registration declares many permissions (Mail, Calendar, Files, Teams, etc.), some organizations may be uncomfortable consenting to all of them.

### Strategy: One App Registration, Multiple Server Instances

You can use a single App Registration with all permissions declared, but run different MCP server instances with different `--preset` or `--enabled-tools` configurations. Each instance only requests the scopes it actually needs during the OAuth flow.

The server dynamically builds its scope list based on which tools are enabled (see `buildScopesFromEndpoints()` in `auth.ts`). The OAuth discovery endpoint (`/.well-known/oauth-authorization-server`) reflects only the scopes for the enabled tools.

**Example configurations:**

```bash
# Mail-only instance — consent screen shows only Mail permissions
ms-365-mcp-server --http 3001 --preset mail

# Calendar-only instance
ms-365-mcp-server --http 3002 --preset calendar

# Mail + Calendar instance
ms-365-mcp-server --http 3003 --preset mail,calendar

# Full suite
ms-365-mcp-server --http 3004 --preset all
```

### Incremental Upgrades

Azure tracks consent per-scope, not per-App Registration. If a customer starts with "Mail only" and later upgrades to "Mail + Calendar":

1. Point them at the instance with the broader preset.
2. They receive a new consent prompt for only the additional calendar scopes.
3. The existing mail consent carries over — no re-consent needed.

## Rate Limiting

Microsoft Graph API rate limits are scoped **per-app, per-tenant** — not globally. One customer's heavy usage does not affect other customers.

Known limits (approximate, Microsoft doesn't publish exact numbers for all endpoints):

| Workload | Approximate limit |
|---|---|
| Outlook (Mail, Calendar, Contacts) | ~10,000 requests per 10 minutes per app per mailbox |
| OneDrive / SharePoint | Varies, generally more restrictive on uploads |
| Teams | ~30-60 requests/second per app per tenant |

When throttled, Microsoft returns `429 Too Many Requests` with a `Retry-After` header. The server currently does not implement automatic retry-on-429 — the error is returned to the MCP client.

There are also app-wide global limits, but these are much higher and only relevant at very large scale (many tenants, all active simultaneously).

## Token Management

In HTTP mode, the server does not manage token storage — your platform handles that externally. The MSAL token cache and keytar integration are only used in stdio mode.

Your platform is responsible for:

- Storing access and refresh tokens securely per customer
- Refreshing access tokens when they expire (~60 minutes)
- Passing the current access token on each MCP request via the `Authorization: Bearer` header

The server does handle one edge case: if an access token expires during a request (401 from Graph API), it will attempt to refresh using the refresh token from the request context and retry. This uses the server's configured `clientId`/`clientSecret`, so the tokens must have been originally obtained against the same App Registration.
