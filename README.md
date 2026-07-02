# mcp-sharepoint-writer

Serveur MCP (Streamable HTTP) qui donne à Claude.ai la capacité de **déposer des fichiers dans une bibliothèque SharePoint** via Microsoft Graph API.

Le connecteur Microsoft 365 officiel d'Anthropic couvre uniquement la lecture. Ce projet comble le manque en exposant un seul outil MCP : `upload_to_sharepoint`.

---

## Architecture

```
Claude.ai
  | HTTPS + Bearer API key
  v
VPS (Nginx Proxy Manager -> Node.js :3000)
  | OAuth 2.0 Client Credentials (MSAL, cache 1h)
  v
Microsoft Graph API  PUT /items/root:/{filename}:/content
  v
Bibliothèque SharePoint cible
```

Le token Microsoft n'est jamais exposé à Claude. Le serveur l'acquiert de façon autonome via MSAL.

---

## Prérequis

- Node.js >= 20
- Un VPS avec HTTPS (reverse proxy type Nginx)
- Une App Registration Azure (Entra ID) avec :
  - Permission applicative `Sites.Selected` (pas `Sites.ReadWrite.All`)
  - Admin consent accordé
  - Accès explicite au site SharePoint cible

---

## Installation

```bash
git clone https://github.com/qurisk/mcp-sharepoint-writer
cd mcp-sharepoint-writer
npm install
cp .env.example .env
# Remplir .env avec les credentials Azure et la clé API
```

---

## Configuration

Copier `.env.example` en `.env` et renseigner les 6 variables :

| Variable | Description |
|---|---|
| `MCP_API_KEY` | Clé aléatoire (ex: `openssl rand -hex 32`) |
| `AZURE_TENANT_ID` | ID du tenant Entra ID |
| `AZURE_CLIENT_ID` | ID de l'App Registration |
| `AZURE_CLIENT_SECRET` | Secret de l'App Registration |
| `SHAREPOINT_SITE_ID` | ID du site SharePoint cible |
| `SHAREPOINT_DRIVE_ID` | ID de la bibliothèque de documents |

Pour obtenir les IDs SharePoint :

```bash
# Site ID
GET https://graph.microsoft.com/v1.0/sites/{hostname}:/{site-path}

# Drive ID
GET https://graph.microsoft.com/v1.0/sites/{site-id}/drives
```

---

## Lancement

```bash
# Développement
npm start

# Production avec pm2
pm2 start 'node --env-file=/chemin/vers/.env index.js' --name mcp-sharepoint-writer
pm2 save
```

---

## Intégration dans Claude.ai

1. Ouvrir Claude.ai > **Settings > Integrations > Add custom integration**
2. URL : `https://votre-domaine/mcp`
3. Type : Bearer token
4. Valeur : contenu de `MCP_API_KEY`

Claude peut ensuite appeler `upload_to_sharepoint(file_content, target_filename)`.

---

## Comportement de l'outil

- **Fichier inexistant** : création directe
- **Fichier existant** : nouvelle version (historique SharePoint préservé, pas de suppression)
- **Auth invalide** : HTTP 401 avant tout traitement
- **Scope** : limité au site SharePoint configuré (Sites.Selected)

---

## Sécurité

- L'authentification Azure se fait côté serveur exclusivement
- `.env` doit être en `chmod 600` et hors du dépôt git
- L'API key MCP est vérifiée sur chaque requête avant tout traitement

---

## Stack

| Composant | Choix |
|---|---|
| Runtime | Node.js 20 (ESM) |
| MCP SDK | `@modelcontextprotocol/sdk` |
| Auth Graph | `@azure/msal-node` |
| Process manager | pm2 |
| Reverse proxy | Nginx Proxy Manager |
