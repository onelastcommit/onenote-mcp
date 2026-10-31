# @onelastcommit/onenote-mcp

[![npm version](https://img.shields.io/npm/v/@onelastcommit/onenote-mcp.svg)](https://www.npmjs.com/package/@onelastcommit/onenote-mcp)
[![CI](https://github.com/onelastcommit/onenote-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/onelastcommit/onenote-mcp/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Part of [One Last Commit](https://github.com/onelastcommit): small ideas, taken further than strictly necessary.

> **This package has moved.** It was previously published as `@atomiclabs97/onenote-mcp` and then `@1percentlabs/onenote-mcp`, and is now `@onelastcommit/onenote-mcp`. If you have an old name in a config file, switch to the new one. The old packages are deprecated and no longer updated. Versions carry on unbroken, so `0.2.1` is the direct successor to `0.2.0`.

An [MCP](https://modelcontextprotocol.io) server for **Microsoft OneNote**. Bring your notebooks into Claude, Cursor, and any MCP-compatible client — list notebooks and sections, full-text search across pages, read individual pages, and create or delete pages from natural language. Authentication uses Microsoft's device-code flow against your own Entra ID app registration, so your data and credentials never leave your machine.

> _Screenshot/demo coming soon — drop a GIF in `docs/images/demo.gif` and reference it here._

---

## Two ways to run it

| You want to use OneNote in… | Run mode | Start here |
| --- | --- | --- |
| **Claude Desktop, Cursor, Claude Code** on your laptop | `stdio` — the server runs as a subprocess of the client; tokens never leave your disk | [Quick start](#quick-start) below |
| **[claude.ai](https://claude.ai) web + iOS/Android**, or any remote client | `http` — you self-host the server and add it as a Connector | [Remote transport](#remote-transport-http) → [Self-hosting](#self-hosting) |

Stdio is the default and needs no hosting. The HTTP transport is opt-in.

---

## Quick start

### 1. Register a Microsoft Entra app

The server talks to Microsoft Graph using a Microsoft Entra (Azure AD) **app registration** that you own. This takes about 2 minutes.

> Screenshots referenced below live in [`docs/images/`](docs/images/). They are placeholders today — contributions welcome.

#### a. Create the registration

1. Go to the [Microsoft Entra admin center → App registrations](https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) and click **+ New registration**.
2. **Name**: anything (e.g. `onenote-mcp`).
3. **Supported account types**: choose **"Accounts in any organizational directory and personal Microsoft accounts"** if you want both work and personal OneNote to work; otherwise pick what matches your tenant.
4. **Redirect URI**: leave blank — the device-code flow doesn't need one.
5. Click **Register**.

   ![Register an application](docs/images/01-register.png)

#### b. Add the API permissions

1. In the new app's left nav, click **API permissions** → **+ Add a permission** → **Microsoft Graph** → **Delegated permissions**.
2. Search for and check both:
   - `Notes.ReadWrite`
   - `offline_access`
3. Click **Add permissions**. (No admin consent is required for personal Microsoft accounts. Work/school tenants may need an admin to grant consent for the directory.)

   ![Add Notes.ReadWrite + offline_access](docs/images/02-permissions.png)

#### c. Allow the public client flow

The device-code flow is a "public client" flow — it doesn't use a client secret.

1. Go to **Authentication** in the left nav.
2. Scroll to **Advanced settings** → **Allow public client flows** and toggle it **Yes**.
3. Click **Save**.

   ![Allow public client flows](docs/images/03-public-client.png)

#### d. Grab the client ID

Back on the app's **Overview** page, copy the **Application (client) ID**. You'll pass it to the server via the `ONENOTE_MCP_CLIENT_ID` environment variable.

  ![Application (client) ID on the Overview page](docs/images/04-client-id.png)

### 2. Sign in

```sh
ONENOTE_MCP_CLIENT_ID=<your-app-client-id> npx @onelastcommit/onenote-mcp login
```

This prints a code and a URL like:

```
To sign in, use a web browser to open https://microsoft.com/devicelogin
and enter the code ABCD-1234 to authenticate.
```

Open the URL, paste the code, sign in with the Microsoft account whose OneNote you want to access, and approve the requested permissions. The refresh token is then cached at `~/.config/onenote-mcp/tokens.json` (mode `600`) so the MCP server can run silently.

To sign out:

```sh
ONENOTE_MCP_CLIENT_ID=<your-app-client-id> npx @onelastcommit/onenote-mcp logout
```

### 3. Wire it into your MCP client

#### Claude Desktop

Edit `~/Library/Application Support/Claude/claude_desktop_config.json` (macOS) or `%APPDATA%\Claude\claude_desktop_config.json` (Windows):

```jsonc
{
  "mcpServers": {
    "onenote": {
      "command": "npx",
      "args": ["-y", "@onelastcommit/onenote-mcp"],
      "env": {
        "ONENOTE_MCP_CLIENT_ID": "your-app-client-id"
      }
    }
  }
}
```

Restart Claude Desktop. The OneNote tools should appear in the tool picker.

#### Cursor

Add to `~/.cursor/mcp.json`:

```jsonc
{
  "mcpServers": {
    "onenote": {
      "command": "npx",
      "args": ["-y", "@onelastcommit/onenote-mcp"],
      "env": {
        "ONENOTE_MCP_CLIENT_ID": "your-app-client-id"
      }
    }
  }
}
```

#### Anything else

Any MCP-compatible client that supports stdio servers will work. Run `npx @onelastcommit/onenote-mcp` with `ONENOTE_MCP_CLIENT_ID` set in the environment.

---

## Remote transport (HTTP)

In addition to the default stdio mode, the server can speak the MCP **Streamable HTTP** transport. Use this when the MCP client lives somewhere other than the same machine as the server — for example, [claude.ai](https://claude.ai) Connectors (web, iOS, Android), a self-hosted bridge on a home server, or a backend integration.

**Stdio is still the default.** Nothing changes for existing Claude Desktop / Cursor users.

### When to use which

| Transport | When |
| --- | --- |
| `stdio` (default) | Local desktop clients (Claude Desktop, Cursor, Claude Code). Tokens never leave your disk; the server runs as a subprocess of the client. |
| `http` | Remote clients (claude.ai web/mobile Connectors), self-hosted deployments, any setup where the server runs separately from the client. |

### Boot it

```sh
# Generate a strong shared secret first — clients must send this as `Authorization: Bearer <token>`.
export ONENOTE_MCP_HTTP_TOKEN="$(openssl rand -base64 32)"

ONENOTE_MCP_CLIENT_ID=<your-app-client-id> \
  npx @onelastcommit/onenote-mcp --transport http --port 3000
```

The server refuses to start without `ONENOTE_MCP_HTTP_TOKEN` — a remote MCP endpoint with no auth is the entire world reading your OneNote.

### Endpoints

| Method | Path | Auth | Purpose |
| --- | --- | --- | --- |
| `POST` / `GET` / `DELETE` | `/mcp` | `Authorization: Bearer <token>` | MCP Streamable HTTP transport. All tools mounted here. |
| `GET` | `/healthz` | none | Liveness probe; returns 200 with no body. For platform health checks. |

Bearer-token comparison is constant-time. Wrong or missing tokens return `401 Unauthorized` with a `WWW-Authenticate: Bearer` header.

### Defaults and overrides

| Flag | Env var | Default | Notes |
| --- | --- | --- | --- |
| `--transport <stdio\|http>` | — | `stdio` | |
| `--host <addr>` | `ONENOTE_MCP_HTTP_HOST` | `127.0.0.1` | Set to `0.0.0.0` to listen on all interfaces (containers / PaaS). |
| `--port <n>` | `ONENOTE_MCP_HTTP_PORT` | `3000` | |

Flags win over env vars.

### Pointing a client at it

```sh
# Sanity check
curl http://127.0.0.1:3000/healthz                                 # → 200
curl -X POST -H "Authorization: Bearer $ONENOTE_MCP_HTTP_TOKEN" \
     -H "Content-Type: application/json" \
     -H "Accept: application/json, text/event-stream" \
     -d '{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}' \
     http://127.0.0.1:3000/mcp
```

Sign in with `ONENOTE_MCP_CLIENT_ID=<your-app-client-id> npx @onelastcommit/onenote-mcp login` (one-time, on the host running the server) before any tools are invoked — the device-code flow is the same as the stdio path and caches tokens at `~/.config/onenote-mcp/tokens.json`. The bearer token is **server↔client auth only**; the Microsoft Graph credentials still come from the cached refresh token.

On a headless host (a container, a PaaS) there's no terminal for the device-code login. Run `login` on your laptop instead, then pass the resulting `tokens.json` to the server via the `ONENOTE_MCP_TOKEN_CACHE` env var — on first boot in HTTP mode the server seeds its cache from it. The [deployment guides](#self-hosting) walk through this.

### Rotating the bearer token

1. Pick a new token: `openssl rand -base64 32`.
2. Update `ONENOTE_MCP_HTTP_TOKEN` and restart the server.
3. Update the token everywhere a client uses it.

No session state is kept between requests, so rotation is just a restart — there is nothing to migrate.

### Multi-tenant?

No. v0.2 assumes **one OneNote account per deployment**. The bearer token gates access to whichever account `onenote-mcp login` was last run for on the host. If you need per-user separation today, run one instance per user.

### Limitations

- The HTTP transport is **stateless**: each request gets a fresh transport, no `Mcp-Session-Id` is issued, and the server does not push notifications. Tools that perform a single request → single response (i.e. all eleven shipped tools) work normally.
- No rate-limiting beyond the bearer-token gate. Put the server behind a reverse proxy / API gateway if you need it.

---

## Self-hosting

To use OneNote from [claude.ai](https://claude.ai) (web + mobile) you host the HTTP transport somewhere with a public HTTPS URL, then add it as a custom Connector. Copy-paste recipes:

| Guide | For |
| --- | --- |
| [Fly.io](docs/deployment/fly.md) | Fastest path — free-tier eligible, ~10 minutes, HTTPS handled for you. |
| [Docker (VPS / home server)](docs/deployment/docker.md) | Any host you control; uses the published `ghcr.io/onelastcommit/onenote-mcp` image. |
| [claude.ai Connector](docs/deployment/claude-ai-connector.md) | Wiring the deployed URL into claude.ai and verifying the tools. |

Released Docker images are pushed to the GitHub Container Registry on every version tag.

---

## Tools

| Tool             | Description                                                     | Key inputs                                       |
| ---------------- | --------------------------------------------------------------- | ------------------------------------------------ |
| `list_notebooks`       | Lists all OneNote notebooks for the signed-in user.             | _(none)_                                         |
| `list_sections`        | Lists sections, optionally scoped to a single notebook.         | `notebookId?`                                    |
| `list_section_groups`  | Lists section groups (folders), optionally scoped to a notebook. | `notebookId?`                                   |
| `search_pages`         | Full-text search across pages (title + content).                | `query`, `limit?`                                |
| `read_page`            | Returns page metadata + content (HTML or Markdown).             | `pageId`, `format?` (`html` \| `markdown`)       |
| `create_notebook`      | Creates a new top-level notebook.                               | `name`                                           |
| `create_section`       | Creates a section inside a notebook or section group.           | `notebookId?` \| `sectionGroupId?`, `name`       |
| `create_section_group` | Creates a section group inside a notebook or another section group. | `notebookId?` \| `sectionGroupId?`, `name`   |
| `create_page`          | Creates a page in a section. Accepts Markdown (default) or HTML, plus optional binary attachments. | `sectionId`, `title`, `content`, `format?`, `attachments?` |
| `update_page`          | Applies edits to a page: append/prepend/insert/replace/delete elements. | `pageId`, `operations[]`                 |
| `delete_page`          | Permanently deletes a page. **Irreversible.**                   | `pageId`                                         |

---

## Configuration

| Env var                    | Required | Description                                                                       |
| -------------------------- | -------- | --------------------------------------------------------------------------------- |
| `ONENOTE_MCP_CLIENT_ID`    | yes      | Application (client) ID of your Microsoft Entra app registration.                  |
| `ONENOTE_MCP_TENANT_ID`    | no       | Tenant ID. Defaults to `common`, which works for both personal and work accounts. |
| `ONENOTE_MCP_HTTP_TOKEN`   | http only | Shared bearer token for the HTTP transport. Required when `--transport http`. |
| `ONENOTE_MCP_HTTP_HOST`    | no       | Bind address for the HTTP transport. Defaults to `127.0.0.1`. `--host` wins.       |
| `ONENOTE_MCP_HTTP_PORT`    | no       | Listen port for the HTTP transport. Defaults to `3000`. `--port` wins.             |
| `ONENOTE_MCP_TOKEN_CACHE`  | no       | Verbatim `tokens.json` contents used to seed the cache on a headless host. Applied only in `--transport http` mode when no cached token exists yet. |
| `XDG_CONFIG_HOME`          | no       | Override the config directory. Tokens are stored at `<dir>/onenote-mcp/tokens.json`. |

---

## Known limitations

- **Attachments are sent in-memory.** `create_page` reads attachment files fully into memory before posting; very large files (~150 MB+) may strain Node's heap. Streamed uploads are a future enhancement.
- **`update_page` targets are raw `data-id` selectors.** To edit a specific element, read the page first and pull the `data-id` attribute out of the returned HTML. Higher-level selectors (e.g. "the section under heading X") are tracked for a follow-up.
- **Search latency.** Microsoft Graph's `$search` against `/me/onenote/pages` can take a few seconds against large notebooks; the server retries on 429s with exponential backoff.

---

## Contributing

PRs welcome. The codebase aims to stay small and focused.

```sh
git clone https://github.com/onelastcommit/onenote-mcp
cd onenote-mcp
yarn install
yarn build
yarn test
```

- `yarn typecheck` — `tsc --noEmit` (covers `src/` + `scripts/`)
- `yarn lint` — ESLint
- `yarn test` — Vitest (unit tests, mocked Graph)
- `yarn dev` — incremental rebuild on save (`tsc --watch` + `tsc-alias --watch`); `dist/` stays runnable

> Use Yarn (Classic), not npm. The repo's lockfile is `yarn.lock` — `package-lock.json` should not be committed.

### End-to-end smoke test

`scripts/smoke.ts` exercises every shipped tool against a real OneNote account. It's not in CI — needs real credentials.

```sh
# One-time: register an Entra app (see Quick start above) and sign in
ONENOTE_MCP_CLIENT_ID=<your-client-id> yarn build
ONENOTE_MCP_CLIENT_ID=<your-client-id> node dist/cli.js login

# Then run the smoke test
ONENOTE_MCP_CLIENT_ID=<your-client-id> yarn smoke
```

It's idempotent: reuses (or creates) a notebook called **OneNote MCP Smoke Test**, walks through 13 steps covering auth, list/search/read/create/update/delete and the multipart attachment path, and cleans up the pages it creates. Sections and section groups are left behind for the next run.

If you're using an AI coding assistant (Claude Code, Cursor, etc.), see [`CLAUDE.md`](CLAUDE.md) for project conventions, the arrow-function rule, and the tool-authoring checklist.

Commits follow [Conventional Commits](https://www.conventionalcommits.org/) (`feat:`, `fix:`, `chore:`, etc.). CI runs typecheck + lint + test + build on every push to `main` and every PR.

### Roadmap

- [x] `update_page` — in-place edits via Graph's `PATCH` syntax
- [x] `create_section` / `create_notebook`
- [x] Image and attachment upload (multipart create_page)
- [x] Section group support (list + create + section-in-group targeting)
- [x] HTTP (Streamable HTTP) transport with bearer-token auth
- [x] Self-hosted deployment guide (Dockerfile + Fly.io recipe + claude.ai Connector validation)
- [ ] Streamed attachment uploads (avoid in-memory buffering for large files)
- [ ] Resource-style page browsing (alongside the tool surface)

---

## License

[MIT](LICENSE) © Ahmad Al Mezaal
