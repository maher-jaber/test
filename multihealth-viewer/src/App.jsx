import React, { useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { PublicClientApplication } from "@azure/msal-browser";
import { Client } from "@microsoft/microsoft-graph-client";
import 'regenerator-runtime/runtime';
import * as microsoftTeams from "@microsoft/teams-js";

const AZURE_APP_ID = "1135fab5-62e8-4cb1-b472-880c477a8812";
const GRAPH_SCOPE = "https://graph.microsoft.com/Files.Read";

const msalConfig = {
  auth: {
    clientId: AZURE_APP_ID,
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin
  }
};

function decodeJwt(token) {
  try {
    return JSON.parse(atob(token.split('.')[1]));
  } catch (e) {
    return null;
  }
}

function App() {
  const [msalInstance] = useState(new PublicClientApplication(msalConfig));
  const [graphClient, setGraphClient] = useState(null);
  const [files, setFiles] = useState([]);
  const [previewUrl, setPreviewUrl] = useState(null);
  const [error, setError] = useState(null);

  const urlParams = new URLSearchParams(window.location.search);
  const siteUrl = urlParams.get("siteUrl") || "";
  const folderPath = urlParams.get("folderPath") || "";

  /** ✅ Initialisation SSO Teams */
  useEffect(() => {
    microsoftTeams.app.initialize().then(() => {
      microsoftTeams.authentication.getAuthToken({
        resources: [AZURE_APP_ID],
        successCallback: async (teamsToken) => {
          console.log("✅ Token Teams obtenu (SSO)");

          const decoded = decodeJwt(teamsToken);
          console.log("👤 Utilisateur :", decoded?.preferred_username);

          const result = await msalInstance.acquireTokenSilent({
            scopes: [GRAPH_SCOPE],
            account: msalInstance.getAllAccounts()[0],
            forceRefresh: true,
          }).catch(async () => {
            return await msalInstance.acquireTokenByAuthorizationCode({
              scopes: [GRAPH_SCOPE],
            });
          });

          const graph = Client.init({
            authProvider: (done) => done(null, result.accessToken),
          });

          setGraphClient(graph);
        },

        failureCallback: (err) => {
          console.error("❌ Erreur Token Teams :", err);
          setError("Erreur SSO Teams : " + err);
        }
      });
    });
  }, []);


  /** ✅ Lister les PDF */
  async function listPdfs() {
    if (!graphClient) return;

    try {
      const hostname = new URL(siteUrl).hostname;
      const site = await graphClient.api(`/sites/${hostname}`).get();

      const drives = await graphClient.api(`/sites/${site.id}/drives`).get();
      const drive = drives.value.find(d => d.name.toLowerCase().includes("document"));

      const response = await graphClient
        .api(`/drives/${drive.id}/root:${folderPath}:/children`)
        .get();

      setFiles(response.value.filter(f => f.file && f.name.endsWith(".pdf")));

    } catch (err) {
      console.error(err);
      setError(err.message);
    }
  }

  /** ✅ Preview PDF */
  async function previewFile(file) {
    try {
      const preview = await graphClient
        .api(`/drives/${file.parentReference.driveId}/items/${file.id}/preview`)
        .post({});

      setPreviewUrl(preview.getUrl);
    } catch (err) {
      setError(err.message);
    }
  }

  return (
    <div style={{ padding: 20, fontFamily: "Segoe UI" }}>
      <h2>📄 MultiHealth — PDF Viewer</h2>

      <button onClick={listPdfs}>📂 Lister les fichiers PDF</button>

      {error && <div style={{ color: "red" }}>{error}</div>}

      <ul>
        {files.map(f => (
          <li key={f.id}>
            {f.name} <button onClick={() => previewFile(f)}>Aperçu</button>
          </li>
        ))}
      </ul>

      {previewUrl && (
        <iframe src={previewUrl} title="preview"
                style={{ width: "100%", height: "80vh", marginTop: 20 }} />
      )}
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
export default App;
