# nextcloud-mcp

MCP server for Nextcloud — file management via WebDAV with Authelia OIDC authentication.

Implements both SSE (`/sse` + `/messages`) and StreamableHTTP (`/mcp`) transports, with RFC 8414 and RFC 9728 OAuth discovery endpoints.

## Tools

| Tool | Description |
|------|-------------|
| `list_files` | List files and folders in a directory |
| `get_file` | Read text file content (returns binary info for non-text files) |
| `get_file_info` | Get file/folder metadata (size, type, modified, etag) |
| `upload_file` | Create or overwrite a text file |
| `create_folder` | Create a new directory |
| `delete_file` | Delete a file or folder |
| `move_file` | Move or rename a file/folder |
| `copy_file` | Copy a file/folder |
| `search_files` | Search by filename (case-insensitive, recursive) |

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `NEXTCLOUD_URL` | — | Base URL of your Nextcloud instance (e.g. `https://nextcloud.example.com`) |
| `NEXTCLOUD_USERNAME` | — | Nextcloud username |
| `NEXTCLOUD_APP_PASSWORD` | — | Nextcloud app password (Settings → Security → App passwords) |
| `MCP_API_KEY` | — | Optional static API key for simple bearer-token auth |
| `OIDC_INTROSPECTION_URL` | `http://authelia:9091/api/oidc/introspection` | Authelia token introspection endpoint |
| `OIDC_CLIENT_ID` | `nextcloud-mcp` | OIDC client ID registered in Authelia |
| `OIDC_CLIENT_SECRET` | — | OIDC client secret |
| `OAUTH_ISSUER` | — | Public Authelia URL (RFC 8414 issuer) |
| `MCP_SERVER_URL` | `https://nextcloud-mcp.${DOMAIN}` | Public URL of this MCP server (RFC 9728) |
| `LOG_LEVEL` | `info` | Log verbosity: `error` \| `warn` \| `info` \| `debug` \| `trace` |
| `DOMAIN` | — | Domain used for Traefik routing (docker-compose only) |
| `NEXTCLOUD_NETWORK` | `nextcloud_nextcloud-internal` | Docker network name for the Nextcloud stack |

## Deployment

### Docker Compose (with Traefik + Authelia)

1. Copy `.env.example` to `.env` and fill in your values.
2. Register a new OIDC client in Authelia with the client ID from `OIDC_CLIENT_ID` and grant type `authorization_code`.
3. Deploy:

```bash
docker compose up -d
```

The server will be available at `https://nextcloud-mcp.${DOMAIN}`.

### Connect to Claude.ai

Add a new MCP connector in Claude.ai using the URL:

```
https://nextcloud-mcp.your-domain.com/sse
```

Authentication is handled automatically via Authelia OIDC.

### Local development

```bash
cp .env.example .env
# edit .env with your values
npm install
NEXTCLOUD_URL=... NEXTCLOUD_USERNAME=... NEXTCLOUD_APP_PASSWORD=... npm run dev -- --http --port 3000
```

## Authelia Client Configuration

```yaml
identity_providers:
  oidc:
    clients:
      - client_id: nextcloud-mcp
        client_name: Nextcloud MCP
        client_secret: '<hashed-secret>'
        public: false
        authorization_policy: one_factor
        redirect_uris:
          - https://claude.ai/oauth/callback
        scopes:
          - openid
          - profile
          - email
        grant_types:
          - authorization_code
          - refresh_token
        response_types:
          - code
        token_endpoint_auth_method: client_secret_basic
```
