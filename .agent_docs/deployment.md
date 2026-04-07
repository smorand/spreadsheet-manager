# Deployment

## Production Environment

**URL**: https://spreadsheet-manager.scm-platform.org
**VPS IP**: 31.97.54.67
**Container**: `spreadsheet-manager` on `proxy-network`
**Port**: 8080 (internal, nginx proxies HTTPS)

## Endpoints

| Endpoint | Auth | Purpose |
|----------|------|---------|
| `/health` | None | Health check |
| `/.well-known/oauth-protected-resource` | None | RFC 9728 metadata |
| `/.well-known/oauth-authorization-server` | None | RFC 8414 metadata |
| `/oauth/register` | None | Dynamic Client Registration (RFC 7591) |
| `/oauth/authorize` | None | OAuth authorization (redirects to Google) |
| `/oauth/callback` | None | Google OAuth callback |
| `/oauth/token` | None | Token exchange and refresh |
| `/` | Bearer token | MCP Streamable HTTP endpoint |

## Credentials

**OAuth credential file**: `scm-pwd-web.json` (Google OAuth Web Application credentials)
- Local dev: `~/.credentials/scm-pwd-web.json`
- VPS: `/app/data/spreadsheet-manager/scm-pwd-web.json`
- Vault backup: `secret/spreadsheet-manager/scm-pwd-web` (field: `credentials`)

**Restoring credentials from Vault** (on VPS):
```bash
docker exec -e VAULT_ADDR=http://127.0.0.1:8200 -e VAULT_TOKEN=B5MJcZJZMum6xam6GKmEW6je vault \
  vault kv get -field=credentials secret/spreadsheet-manager/scm-pwd-web > /app/data/spreadsheet-manager/scm-pwd-web.json
chmod 600 /app/data/spreadsheet-manager/scm-pwd-web.json
```

## DNS

**Zone**: `scm-platform-org` (Google Cloud DNS)
**Record**: `spreadsheet-manager.scm-platform.org` A `31.97.54.67` (TTL 3600)

```bash
# Create (already done)
gcloud dns record-sets create "spreadsheet-manager.scm-platform.org." \
  --zone=scm-platform-org --type=A --ttl=3600 --rrdatas="31.97.54.67"

# Verify
dig +short spreadsheet-manager.scm-platform.org A
```

## SSL Certificate

**Provider**: Let's Encrypt (auto-renewal via certbot systemd timer)
**Expires**: 2026-07-06

```bash
# Renew manually if needed (on VPS)
cd /opt/nginx-reverse-proxy
LETSENCRYPT_EMAIL=seb.morand@gmail.com ./scripts/letsencrypt.sh renew
./scripts/proxy.sh reload
```

## Deploying a New Version

```bash
# 1. Commit and tag
git tag v1.x.0
git push origin main --tags

# 2. On VPS: re-clone and rebuild
ssh root@31.97.54.67
cd /app/services
rm -rf spreadsheet-manager
git clone --branch v1.x.0 --depth 1 https://github.com/smorand/spreadsheet-manager.git
cd spreadsheet-manager
docker compose -f docker-compose.prod.yml up -d --build

# Data in /app/data/spreadsheet-manager/ is preserved across deployments
```

## VPS File Layout

```
/app/services/spreadsheet-manager/     # Cloned repo
/app/data/spreadsheet-manager/         # Persistent data
  .env                                 # Environment variables
  scm-pwd-web.json                     # OAuth credentials
/logs/spreadsheet-manager/             # Logs
```

## Nginx Route

```
spreadsheet-manager.scm-platform.org -> spreadsheet-manager:8080
```

**Important**: The upstream must point to the container name `spreadsheet-manager`, not `host.docker.internal`.

```bash
# Verify route (on VPS)
cd /opt/nginx-reverse-proxy
./scripts/proxy.sh list | grep spreadsheet

# Fix if needed
./scripts/proxy.sh remove spreadsheet-manager.scm-platform.org
./scripts/proxy.sh add spreadsheet-manager.scm-platform.org 8080 spreadsheet-manager
./scripts/proxy.sh reload
```

## Troubleshooting

**502 Bad Gateway**: nginx upstream points to wrong target. Check route with `proxy.sh list`, fix with commands above.

**Container unhealthy**: Check logs with `docker logs spreadsheet-manager`. Verify credential file exists at `/app/data/spreadsheet-manager/scm-pwd-web.json`.

**OAuth errors**: Verify `BASE_URL` in `.env` matches the actual domain. Verify `scm-pwd-web.json` contains valid Google OAuth Web Application credentials with the redirect URI `https://spreadsheet-manager.scm-platform.org/oauth/callback`.
