import React, { useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { Client } from "@microsoft/microsoft-graph-client";
import 'regenerator-runtime/runtime';
import * as microsoftTeams from "@microsoft/teams-js";
import * as msal from "@azure/msal-browser";

const AZURE_APP_ID = "1135fab5-62e8-4cb1-b472-880c477a8812";


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
    redirectUri: window.location.origin + '/'
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
  const [graphClient, setGraphClient] = useState(null);
  const [files, setFiles] = useState([]);
  const [previewUrl, setPreviewUrl] = useState(null);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);
  const [authStatus, setAuthStatus] = useState("initializing");
  const [account, setAccount] = useState(null);
  const [client, setClient] = useState(null);

  const urlParams = new URLSearchParams(window.location.search);
  const siteUrl = urlParams.get("siteUrl") || "";
  const folderPath = urlParams.get("folderPath") || "";

  /** ✅ Initialisation SSO Teams */
  useEffect(() => {
    async function initializeTeams() {
      try {


        console.log("🔄 Initialisation Teams…");
        await microsoftTeams.app.initialize();
        const context = await microsoftTeams.app.getContext();
        const isDesktop = context.app.host.clientType === "desktop";
        console.log("💻 Mode :", isDesktop ? "Desktop" : "Web");

        if (isDesktop) {
          console.log("🔐 Desktop → Tentative SSO sans popup");

          const authToken = await microsoftTeams.authentication.getAuthToken({
            resources: ["https://graph.microsoft.com"]
          });

          console.log("✅ SSO Desktop OK");
          initGraphClient(authToken);
          setAccount({ username: decodeJwt(authToken)?.preferred_username });
          setAuthStatus("authenticated");
        }


        setTimeout(() => openTeamsAuthDialog(), 300);
        setAuthStatus("waiting_for_web_popup");


      } catch (err) {
        console.error("❌ Erreur SSO Teams:", err);
        openTeamsAuthDialog();

      }
    }

    initializeTeams();
    setLoading(true);
    setTimeout(function () {
      listPdfs();
    }, 6000);

  }, []);


  function openTeamsAuthDialog() {
    microsoftTeams.authentication.authenticate({
      url: window.location.origin + "/auth.html",
      width: 600,
      height: 600,
      successCallback: (accessToken) => {
        console.log("✅ Token reçu depuis auth.html:", accessToken);

        // ✅ Pas besoin de MSAL ici ! On utilise directement le token.
        initGraphClient(accessToken);

        // ✅ Sauvegarder "visuellement" que l'utilisateur est connecté
        setAccount({
          username: decodeJwt(accessToken)?.preferred_username,
          token: accessToken
        });

        setAuthStatus("authenticated");
      },
      failureCallback: (reason) => {
        console.error("❌ Auth dialog erreur:", reason);
        setError(reason);
      }
    });
  }

  function initGraphClient(accessToken) {
    const graph = Client.init({
      authProvider: (done) => done(null, accessToken)
    });

    setClient(graph);
    setGraphClient(graph);
  }

 

  async function listPdfs() {
    setLoading(true);
    setError(null);

    try {
      const hostname = new URL(siteUrl).hostname;
      const pathParts = new URL(siteUrl).pathname.split("/").filter(Boolean);
      const sitePath = pathParts.slice(1).join("/");

      const site = await client.api(`/sites/${hostname}:/sites/${sitePath}`).get();
      const drives = await client.api(`/sites/${site.id}/drives`).get();

      let driveId = drives.value.find(d => d.name.toLowerCase().includes("document"))?.id;
      if (!driveId) throw new Error("❌ Aucune bibliothèque trouvée");

      const response = await client
        .api(`/drives/${driveId}/root:${folderPath}:/children`)
        .get();

      const pdf = response.value.find(f => f.file && f.name.endsWith(".pdf"));

      if (!pdf) throw new Error("❌ Aucun PDF trouvé dans ce dossier");

      setFiles([pdf]);     // stockage si tu veux afficher le nom
      await previewFile(pdf); // ⬅️ affichage immédiat

    } catch (e) {
      setError(e.message);
    }

    setLoading(false);
  }



  /** ✅ Preview PDF avec URL directe SharePoint */
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
    <div style={{ padding: 20, fontFamily: "'Segoe UI', sans-serif", backgroundColor: "#f3f2f1", minHeight: "100vh" }}>
      <h2 style={{ marginBottom: 20, color: "#323130" }}>📄 MultiHealth — PDF Viewer</h2>




      {/* Connexion */}
      {!account && (
        <button
          onClick={openTeamsAuthDialog}
          style={{
            padding: "10px 20px",
            backgroundColor: "#0078d4",
            color: "#ffffff",
            border: "none",
            borderRadius: 6,
            cursor: "pointer",
            marginBottom: 20,
            fontWeight: 500,
            boxShadow: "0 2px 4px rgba(0,0,0,0.1)",
            transition: "background 0.2s"
          }}
          onMouseOver={e => e.currentTarget.style.backgroundColor = "#005a9e"}
          onMouseOut={e => e.currentTarget.style.backgroundColor = "#0078d4"}
        >
          🔐 Se connecter à Microsoft Graph
        </button>
      )}
      {loading && (
        <div style={{ marginBottom: 10 }}>

          ⏳ Chargement...

        </div>
      )}
      {/* Erreur */}
      {error && (
        <div style={{
          color: "#a80000",
          backgroundColor: "#fde7e9",
          padding: 12,
          borderRadius: 6,
          marginTop: 10,
          border: "1px solid #f5c2c7",
          fontWeight: 500
        }}>
          ❌ {error}
        </div>
      )}

      {/* Initialisation */}
      {!graphClient && !error && (
        <div style={{
          color: "#605e5c",
          padding: 10,
          marginTop: 10,
          fontStyle: "italic"
        }}>
          🔄 {authStatus === "teams_initialized" ?
            "Authentification avec ressource personnalisée..." :
            "Initialisation de Teams..."}
        </div>
      )}



      {/* Aperçu PDF */}
      {previewUrl && (
        <div style={{ marginTop: 20 }}>
          <div style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            marginBottom: 10
          }}>
            <h3 style={{ color: "#323130" }}>👁️ Aperçu PDF</h3>

          </div>
          <iframe
            src={previewUrl}
            title="preview"
            style={{
              width: "100%",
              height: "80vh",
              border: "1px solid #ddd",
              borderRadius: 8,
              backgroundColor: "#ffffff",
              boxShadow: "0 1px 5px rgba(0,0,0,0.1)"
            }}
          />
        </div>
      )}
    </div>
  );


}

createRoot(document.getElementById("root")).render(<App />);
export default App;