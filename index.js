import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { createServer } from "node:http";
import { randomBytes } from "node:crypto";
import { z } from "zod";

const {
  MCP_API_KEY,
  OAUTH_CLIENT_ID,
  OAUTH_CLIENT_SECRET,
  BASE_URL,
  AZURE_TENANT_ID,
  AZURE_CLIENT_ID,
  AZURE_CLIENT_SECRET,
  SHAREPOINT_SITE_ID,
  SHAREPOINT_DRIVE_ID,
} = process.env;

// ---------------------------------------------------------------------------
// Azure / SharePoint
// ---------------------------------------------------------------------------
const msalApp = new ConfidentialClientApplication({
  auth: {
    clientId: AZURE_CLIENT_ID,
    clientSecret: AZURE_CLIENT_SECRET,
    authority: `https://login.microsoftonline.com/${AZURE_TENANT_ID}`,
  },
});

async function getToken() {
  const result = await msalApp.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  return result.accessToken;
}

// ---------------------------------------------------------------------------
// MCP server factory
// ---------------------------------------------------------------------------
function createMcpServer() {
  const server = new McpServer({ name: "sharepoint-writer", version: "1.0.0" });

  server.tool(
    "upload_to_sharepoint",
    "Dépose un fichier dans la bibliothèque SharePoint configurée",
    {
      file_content: z.string().describe("Contenu texte du fichier"),
      target_filename: z.string().describe("Nom du fichier cible, ex: rapport.csv"),
    },
    async ({ file_content, target_filename }) => {
      const token = await getToken();
      const res = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_SITE_ID}/drives/${SHAREPOINT_DRIVE_ID}/items/root:/${encodeURIComponent(target_filename)}:/content`,
        {
          method: "PUT",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "text/plain; charset=utf-8",
          },
          body: file_content,
        }
      );
      const data = await res.json();
      if (!res.ok) {
        return { content: [{ type: "text", text: `Erreur ${res.status}: ${data.error?.message}` }] };
      }
      return { content: [{ type: "text", text: `Fichier "${data.name}" déposé. URL: ${data.webUrl}` }] };
    }
  );

  return server;
}

// ---------------------------------------------------------------------------
// OAuth 2.0 - Authorization Code flow (minimal, auto-approve)
// codes: Map<code, { clientId, redirectUri, expiresAt }>
// ---------------------------------------------------------------------------
const pendingCodes = new Map();

function oauthMeta() {
  return {
    issuer: BASE_URL,
    authorization_endpoint: `${BASE_URL}/oauth/authorize`,
    token_endpoint: `${BASE_URL}/oauth/token`,
    response_types_supported: ["code"],
    grant_types_supported: ["authorization_code"],
    code_challenge_methods_supported: [],
  };
}

function handleAuthorize(req, res) {
  const url = new URL(req.url, BASE_URL);
  const clientId    = url.searchParams.get("client_id");
  const redirectUri = url.searchParams.get("redirect_uri");
  const state       = url.searchParams.get("state") ?? "";

  if (clientId !== OAUTH_CLIENT_ID) {
    res.writeHead(400).end(JSON.stringify({ error: "invalid_client" }));
    return;
  }

  const code = randomBytes(24).toString("hex");
  pendingCodes.set(code, {
    clientId,
    redirectUri,
    expiresAt: Date.now() + 5 * 60 * 1000,
  });

  const target = new URL(redirectUri);
  target.searchParams.set("code", code);
  if (state) target.searchParams.set("state", state);

  res.writeHead(302, { Location: target.toString() }).end();
}

async function handleToken(req, res, body) {
  const params = new URLSearchParams(body);
  const grantType    = params.get("grant_type");
  const code         = params.get("code");
  const clientId     = params.get("client_id");
  const clientSecret = params.get("client_secret");

  // Also accept Basic auth
  const authHeader = req.headers.authorization ?? "";
  if (authHeader.startsWith("Basic ")) {
    const decoded = Buffer.from(authHeader.slice(6), "base64").toString();
    const [id, secret] = decoded.split(":");
    if (!clientId) params.set("client_id", id);
    if (!clientSecret) params.set("client_secret", secret);
  }

  const cid = clientId || params.get("client_id");
  const cs  = clientSecret || params.get("client_secret");

  if (cid !== OAUTH_CLIENT_ID || cs !== OAUTH_CLIENT_SECRET) {
    res.writeHead(401, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: "invalid_client" }));
    return;
  }

  if (grantType !== "authorization_code") {
    res.writeHead(400, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: "unsupported_grant_type" }));
    return;
  }

  const entry = pendingCodes.get(code);
  if (!entry || Date.now() > entry.expiresAt) {
    res.writeHead(400, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: "invalid_grant" }));
    return;
  }
  pendingCodes.delete(code);

  res.writeHead(200, { "Content-Type": "application/json" });
  res.end(JSON.stringify({
    access_token: MCP_API_KEY,
    token_type: "Bearer",
    expires_in: 31536000,
  }));
}

// ---------------------------------------------------------------------------
// HTTP server
// ---------------------------------------------------------------------------
const httpServer = createServer(async (req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Authorization, Content-Type, Accept");
  if (req.method === "OPTIONS") { res.writeHead(204).end(); return; }

  const path = new URL(req.url, BASE_URL).pathname;

  // OAuth metadata
  if (path === "/.well-known/oauth-authorization-server") {
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify(oauthMeta()));
    return;
  }

  // Authorization endpoint
  if (path === "/oauth/authorize") {
    handleAuthorize(req, res);
    return;
  }

  // Token endpoint
  if (path === "/oauth/token" && req.method === "POST") {
    let body = "";
    req.on("data", (c) => (body += c));
    req.on("end", () => handleToken(req, res, body));
    return;
  }

  // MCP endpoint
  if (path === "/mcp") {
    const auth = req.headers.authorization ?? "";
    if (auth !== `Bearer ${MCP_API_KEY}`) {
      res.writeHead(401, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ error: "Unauthorized" }));
      return;
    }
    let body = "";
    req.on("data", (c) => (body += c));
    req.on("end", async () => {
      try {
        const server    = createMcpServer();
        const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
        await server.connect(transport);
        await transport.handleRequest(req, res, body ? JSON.parse(body) : undefined);
      } catch (err) {
        console.error(err);
        if (!res.headersSent) res.writeHead(500).end(JSON.stringify({ error: "Internal error" }));
      }
    });
    return;
  }

  res.writeHead(404).end();
});

httpServer.listen(3000, () => console.log("MCP SharePoint server démarré sur :3000"));
