import React, { useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { Client } from "@microsoft/microsoft-graph-client";
import 'regenerator-runtime/runtime';
import * as microsoftTeams from "@microsoft/teams-js";
import * as msal from "@azure/msal-browser";

function decodeJwt(token) {
  try {
    return JSON.parse(atob(token.split('.')[1]));
  } catch (e) {
    return null;
  }
}

const msalConfig = {
  auth: {
    clientId: process.env.REACT_APP_CLIENT_ID || "",
    authority: `https://login.microsoftonline.com/${process.env.REACT_APP_TENANT_ID || "common"}`,
    redirectUri: window.location.origin + '/'  // <-- redirect principal
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false
  }
};

const loginRequest = {
  scopes: ["openid", "profile", "Files.Read.All", "Sites.Read.All", "offline_access", "User.Read"]
};

function App() {

  const [msalInstance] = useState(new msal.PublicClientApplication(msalConfig));
  const [client, setClient] = useState(null);
  const [files, setFiles] = useState([]);
  const [previewUrl, setPreviewUrl] = useState(null);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);
  const [authStatus, setAuthStatus] = useState("initializing");
  const [account, setAccount] = useState(null);

  const urlParams = new URLSearchParams(window.location.search);
  const siteUrl = urlParams.get("siteUrl") || "";
  const folderPath = urlParams.get("folderPath") || "";

  /** ✅ Initialisation SSO Teams */
  useEffect(() => {
    const initializeTeams = async () => {
      try {
        console.log("🔄 Initialisation Teams...");
        await microsoftTeams.app.initialize();
        console.log("✅ Teams initialisé");

        setAuthStatus("teams_initialized");

        const authToken = await microsoftTeams.authentication.getAuthToken({
          resources: ["https://graph.microsoft.com"]
        });

        console.log("✅ Token SSO reçu");
        console.log("👤 User:", decodeJwt(authToken)?.preferred_username);

        const graphClient = Client.init({
          authProvider: (done) => done(null, authToken)
        });

        setClient(graphClient);
        setAuthStatus("authenticated");

      } catch (err) {
        console.warn("⚠️ SSO impossible → on bascule sur Auth dialog");
        setAuthStatus("error");
      }
    };

    initializeTeams();
  }, []);

  /** ✅ Auth dans dialog Teams */
  function openTeamsAuthDialog() {
    microsoftTeams.authentication.authenticate({
      url: window.location.origin + "/auth.html",
      width: 600,
      height: 600,
      successCallback: async (accessToken) => {
        console.log("✅ Token reçu depuis auth.html:", accessToken);

        const account = msalInstance.getAllAccounts()[0];
        setAccount(account);

        initGraphClient(account, accessToken);
      },
      failureCallback: (reason) => {
        console.error("❌ Auth dialog erreur:", reason);
        setError(reason);
      }
    });
  }

  /** ✅ Création du client Graph */
  function initGraphClient(activeAccount, cachedToken = null) {
    const graphClient = Client.init({
      authProvider: async (done) => {
        try {
          if (cachedToken) {
            return done(null, cachedToken);
          }

          const response = await msalInstance.acquireTokenSilent({
            ...loginRequest,
            account: activeAccount
          });

          done(null, response.accessToken);
        } catch (err) {
          console.error("⚠️ acquireTokenSilent failed");
          done(err, null);
        }
      }
    });

    setClient(graphClient);
  }

  /** ✅ Tester accès à Graph */
  async function testGraphConnection() {
    if (!client) return;
    try {
      setLoading(true);
      const user = await client.api('/me').get();
      console.log("✅ Graph OK:", user.displayName);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }

  /** ✅ Lister les PDFs */
  async function listPdfs() {
    if (!client) return;

    setLoading(true);
    setError(null);

    try {
      const hostname = new URL(siteUrl).hostname;
      const pathParts = new URL(siteUrl).pathname.split("/").filter(Boolean);
      const sitePath = pathParts.slice(1).join("/");

      const site = await client.api(`/sites/${hostname}:/sites/${sitePath}`).get();
      const drives = await client.api(`/sites/${site.id}/drives`).get();

      const drive = drives.value.find(d => d.name.toLowerCase().includes("document"));
      if (!drive) throw new Error("Aucune bibliothèque trouvée");

      const response = await client.api(`/drives/${drive.id}/root:${folderPath}:/children`).get();
      setFiles(response.value.filter(f => f.file && f.name.endsWith(".pdf")));

    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }

  /** ✅ Preview PDF */
  async function previewFile(item) {
    try {
      const res = await client.api(`/drives/${item.parentReference.driveId}/items/${item.id}/preview`).post({});
      setPreviewUrl(res.getUrl);
    } catch (err) {
      setError(err.message);
    }
  }

  return (
    <div style={{ padding: 20, fontFamily: "Segoe UI, sans-serif" }}>
      <h2>📄 MultiHealth — PDF Viewer</h2>

      <p><strong>Statut:</strong> {authStatus}</p>

      {!client && (
        <button onClick={openTeamsAuthDialog} style={{ padding: 10, background: "#0078d4", color: "white" }}>
          🔐 Se connecter à Microsoft Graph
        </button>
      )}

      {client && (
        <>
          <button onClick={listPdfs} style={{ padding: 10, marginRight: 10 }}>📂 Lister les PDF</button>
          <button onClick={testGraphConnection} style={{ padding: 10 }}>🧪 Test Graph</button>
        </>
      )}

      {error && <div style={{ color: "red" }}>❌ {error}</div>}

      {files.length > 0 && (
        <ul>
          {files.map(f => (
            <li key={f.id}>
              {f.name}
              <button onClick={() => previewFile(f)}>Preview</button>
            </li>
          ))}
        </ul>
      )}

      {previewUrl && (
        <iframe src={previewUrl} style={{ width: "100%", height: "80vh" }} />
      )}
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
export default App;
