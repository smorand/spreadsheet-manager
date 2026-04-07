# Deployment

## Production Environment

**URL**: https://spreadsheet-manager.scm-platform.org
**MCP endpoint**: https://spreadsheet-manager.scm-platform.org/mcp
**VPS IP**: 31.97.54.67
**Container**: `spreadsheet-manager` on `proxy-network`
**Port**: 8080 (internal, nginx proxies HTTPS)

## Endpoints

| Endpoint | Auth | Purpose |
|----------|------|---------|
| `/health` | None | Health check |
| `/mcp` | Bearer token | MCP Streamable HTTP endpoint |
| `/.well-known/oauth-protected-resource` | None | RFC 9728 metadata |
| `/.well-known/oauth-authorization-server` | None | RFC 8414 metadata |
| `/oauth/register` | None | Dynamic Client Registration (RFC 7591) |
| `/oauth/authorize` | None | OAuth authorization (redirects to Google) |
| `/oauth/callback` | None | Google OAuth callback |
| `/oauth/token` | None | Token exchange and refresh |

## Credentials

**Google OAuth Web Application credentials** (`scm-pwd-web.json`) are stored in HashiCorp Vault on the VPS. This is a shared credential used by multiple MCP servers.

**Credential loading priority**:
1. GCP Secret Manager (for Cloud Run, if SECRET_PROJECT + SECRET_NAME are set)
2. HashiCorp Vault (for VPS, if VAULT_ADDR + VAULT_SECRET_PATH are set)
3. Local file (for development, via CREDENTIAL_FILE or `~/.credentials/scm-pwd-web.json`)

**Vault configuration** (in `/app/data/spreadsheet-manager/.env`):
```env
VAULT_ADDR=http://vault:8200
VAULT_TOKEN=<vault-token>
VAULT_SECRET_PATH=secret/credentials/google-credentials
```

**Vault secret path**: `secret/credentials/google-credentials` (shared, field: `credentials`)

**Managing the shared credential in Vault**:
```bash
# Read
docker exec -e VAULT_ADDR=http://127.0.0.1:8200 -e VAULT_TOKEN=<token> vault \
  vault kv get -field=credentials secret/credentials/google-credentials

# Write (from local file)
docker exec -e VAULT_ADDR=http://127.0.0.1:8200 -e VAULT_TOKEN=<token> vault \
  vault kv put secret/credentials/google-credentials credentials="$(cat scm-pwd-web.json)"
```

## DNS

**Zone**: `scm-platform-org` (Google Cloud DNS)
**Record**: `spreadsheet-manager.scm-platform.org` A `31.97.54.67` (TTL 3600)

```bash
# Create (already done)
gcloud dns record-sets create "spreadsheet-manager.scm-platform.org." \
  --zone=scm-platform-org --type=A --ttl=3600 --rrdatas="31.97.54.67"
```

## SSL Certificate

**Provider**: Let's Encrypt (auto-renewal via certbot systemd timer)
**Obtained automatically** by `vps-deploy.sh` when `LETSENCRYPT_EMAIL` is set.

## Deploying / Updating

```bash
# 1. Commit and tag
git tag v1.x.0
git push origin main --tags

# 2. Deploy via vps-deploy.sh (on VPS)
ssh root@31.97.54.67
cd /opt/nginx-reverse-proxy
./scripts/vps-undeploy.sh spreadsheet-manager
LETSENCRYPT_EMAIL=seb.morand@gmail.com ./scripts/vps-deploy.sh smorand/spreadsheet-manager@v1.x.0 prod spreadsheet-manager.scm-platform.org:8080 ./environments

# Data in /app/data/spreadsheet-manager/ (.env, Vault config) is preserved across deployments
```

## VPS File Layout

```
/app/services/spreadsheet-manager/     # Cloned repo (rebuilt on each deploy)
/app/data/spreadsheet-manager/         # Persistent data (preserved)
  .env                                 # Environment variables (Vault config)
/logs/spreadsheet-manager/             # Logs
```

No credential files on disk. Credentials are loaded from Vault at runtime.

## Troubleshooting

**502 Bad Gateway**: Nginx upstream wrong. Verify with `proxy.sh list | grep spreadsheet`. The upstream should be `spreadsheet-manager`, not `host.docker.internal`. This is fixed in `vps-deploy.sh` (passes container name as upstream).

**OAuth configuration error (500)**: Vault unreachable or wrong secret path. Check `docker logs spreadsheet-manager` for Vault errors. Verify env vars with `docker exec spreadsheet-manager env | grep VAULT`.

**Container unhealthy**: Check `docker logs spreadsheet-manager`. Verify the container can reach Vault: `docker exec spreadsheet-manager wget -qO- http://vault:8200/v1/sys/health`.
