
import React, { useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as msal from "@azure/msal-browser";
import { Client } from "@microsoft/microsoft-graph-client";
import 'regenerator-runtime/runtime';
import * as microsoftTeams from "@microsoft/teams-js";

const msalConfig = {
  auth: {
    clientId: process.env.REACT_APP_CLIENT_ID || "",
    authority: `https://login.microsoftonline.com/${process.env.REACT_APP_TENANT_ID || "common"}`,
    redirectUri: window.location.origin + '/'
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false
  }
};

const loginRequest = {
  scopes: ["openid","profile","Files.Read.All","Sites.Read.All","offline_access"]
};
function decodeJwt (token) {
  const base64Url = token.split('.')[1];
  const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
  return JSON.parse(window.atob(base64));
}
function App() {
  const [msalInstance] = useState(new msal.PublicClientApplication(msalConfig));
  const [account, setAccount] = useState(null);
  const [client, setClient] = useState(null);
  const [files, setFiles] = useState([]);
  const [previewUrl, setPreviewUrl] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  const urlParams = new URLSearchParams(window.location.search);
  const siteUrl = urlParams.get('siteUrl') || "";
  const folderPath = urlParams.get('folderPath') || ""; // e.g. /Shared Documents/PDFs

  useEffect(() => {
    microsoftTeams.app.initialize().then(() => {
      microsoftTeams.authentication.getAuthToken({
        successCallback: (token) => {
          console.log("✅ Token récupéré depuis Teams (SSO)");
  
          const decoded = decodeJwt(token);
          setAccount({ username: decoded.preferred_username });
  
          initGraphClient({ token });
        },
        failureCallback: (err) => {
          console.error("❌ Erreur récupération token Teams:", err);
        }
      });
    });
  }, []);
  
  

  function initGraphClient(activeAccount) {
    const msalProvider = {
      getAccessToken: async () => {
        const request = { ...loginRequest, account: activeAccount };
        const response = await msalInstance.acquireTokenSilent(request).catch(async (e) => {
          await msalInstance.loginRedirect(request);
        });
        return response.accessToken;
      }
    };
    const graphClient = Client.init({
      authProvider: async (done) => {
        try {
          const token = await msalProvider.getAccessToken();
          done(null, token);
        } catch (err) {
          done(err, null);
        }
      }
    });
    setClient(graphClient);
  }

  async function signIn() {
    try {
      setError(null);
  
      await msalInstance.loginRedirect(loginRequest);
  
    } catch (err) {
      setError(err.message || String(err));
    }
  }

  async function listPdfs() {
    if (!client) return;
  
    setLoading(true);
    setError(null);
  
    try {
      const hostname = new URL(siteUrl).hostname;
      const pathParts = new URL(siteUrl).pathname.split("/").filter(Boolean);
      const sitePath = pathParts.slice(1).join("/");
  
      console.log("🔍 SITE TARGET:", hostname, sitePath);
  
      // 1️⃣ Récupérer le site
      const site = await client.api(`/sites/${hostname}:/sites/${sitePath}`).get();
      console.log("✅ Site ID:", site.id);
  
      // 2️⃣ Récupérer TOUTES les drives (bibliothèques documentaires)
      const drives = await client.api(`/sites/${site.id}/drives`).get();
      console.log("📂 Drives trouvés:", drives.value.map(d => d.name));
  
      // 3️⃣ Trouver la drive qui contient ton dossier "Administratif"
      let driveId = null;
      for (let d of drives.value) {
        if (d.name.toLowerCase().includes("document")) {
          driveId = d.id;
          console.log("✅ Drive détectée:", d.name, d.id);
          break;
        }
      }
  
      if (!driveId) throw new Error("❌ Aucune bibliothèque de documents trouvée.");
  
      // 4️⃣ Tester l'accès au dossier demandé
      console.log(`🔎 Test: /drives/${driveId}/root:${folderPath}:/children`);
  
      const response = await client
        .api(`/drives/${driveId}/root:${folderPath}:/children`)
        .get();
  
      console.log("✅ Résultat Graph:", response);
  
      const pdfs = response.value.filter(f => f.file && f.name.endsWith(".pdf"));
      setFiles(pdfs);
  
    } catch (e) {
      console.error("❌ ERREUR:", e);
      setError(e.message);
    }
  
    setLoading(false);
  }
  
  async function previewFile(item) {
    if (!client) { setError('Graph client not initialized'); return; }
    setError(null);
    try {
      const res = await client.api(`/drives/${item.parentReference.driveId}/items/${item.id}/preview`).post({});
      if (res && res.getUrl) {
        setPreviewUrl(res.getUrl);
      } else {
        setError('Impossible d\'obtenir l\'URL de preview');
      }
    } catch (err) {
      setError(err.message || String(err));
    }
  }

  return (
    <div style={{ fontFamily: 'Segoe UI, Arial', padding: 20 }}>
      <h2>MultiHealth — PDF Viewer</h2>
     
      <div style={{ marginTop: 12 }}>
        <button onClick={listPdfs} disabled={!account || loading}>Lister les PDFs</button>
      </div>

      {loading && <div>Chargement...</div>}
      {error && <div style={{ color: 'red' }}>{error}</div>}

      <div style={{ display: 'flex', marginTop: 16 }}>
        <div style={{ flex: 1, maxWidth: 360 }}>
          <h4>Fichiers PDF</h4>
          <ul>
            {files.map(f => (
              <li key={f.id} style={{ marginBottom: 8 }}>
                <div><strong>{f.name}</strong></div>
                <div style={{ fontSize: 12 }}>{f.size} bytes</div>
                <div><button onClick={() => previewFile(f)}>Aperçu</button></div>
              </li>
            ))}
            {files.length === 0 && <li>Aucun PDF trouvé (après avoir cliqué sur Lister)</li>}
          </ul>
        </div>

        <div style={{ flex: 3, marginLeft: 16 }}>
          <h4>Aperçu</h4>
          {previewUrl ? (
            <iframe src={previewUrl} title="preview" style={{ width: '100%', height: '80vh', border: '1px solid #ddd' }} />
          ) : (
            <div style={{ color: '#666' }}>Sélectionne un fichier pour voir l'aperçu</div>
          )}
        </div>
      </div>
    </div>
  );
}

const container = document.getElementById('root');
const root = createRoot(container);
root.render(<App />);

export default App;
