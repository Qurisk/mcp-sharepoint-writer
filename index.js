import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { createServer } from "node:http";
import { z } from "zod";

const {
  MCP_API_KEY,
  AZURE_TENANT_ID,
  AZURE_CLIENT_ID,
  AZURE_CLIENT_SECRET,
  SHAREPOINT_SITE_ID,
  SHAREPOINT_DRIVE_ID,
} = process.env;

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

function createMcpServer() {
  const server = new McpServer({ name: "sharepoint-writer", version: "1.0.0" });

  server.tool(
    "upload_to_sharepoint",
    "Dépose un fichier dans la bibliothèque SharePoint configurée",
    {
      file_content: z.string().describe("Contenu texte du fichier"),
      target_filename: z
        .string()
        .describe("Nom du fichier cible, ex: rapport.csv"),
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
        return {
          content: [
            {
              type: "text",
              text: `Erreur ${res.status}: ${data.error?.message}`,
            },
          ],
        };
      }
      return {
        content: [
          {
            type: "text",
            text: `Fichier "${data.name}" déposé. URL: ${data.webUrl}`,
          },
        ],
      };
    }
  );

  return server;
}

const httpServer = createServer(async (req, res) => {
  if (req.headers.authorization !== `Bearer ${MCP_API_KEY}`) {
    res.writeHead(401, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: "Unauthorized" }));
    return;
  }

  let body = "";
  req.on("data", (chunk) => (body += chunk));
  req.on("end", async () => {
    try {
      const server = createMcpServer();
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: undefined,
      });
      await server.connect(transport);
      await transport.handleRequest(
        req,
        res,
        body ? JSON.parse(body) : undefined
      );
    } catch (err) {
      console.error(err);
      if (!res.headersSent) {
        res.writeHead(500).end(JSON.stringify({ error: "Internal error" }));
      }
    }
  });
});

httpServer.listen(3000, () =>
  console.log("MCP SharePoint server démarré sur :3000")
);
