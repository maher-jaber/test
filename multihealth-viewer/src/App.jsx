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
    const initializeTeams = async () => {
      try {
        console.log("🔄 Initialisation Teams...");
        await microsoftTeams.app.initialize();
        console.log("✅ Teams initialisé");
        setAuthStatus("teams_initialized");
        
        // Utiliser la ressource personnalisée
        console.log("🔑 Demande de token pour:");
        const authToken = await microsoftTeams.authentication.getAuthToken({
          resources: ["https://graph.microsoft.com"]
        });
        
        console.log("✅ Token obtenu avec ressource personnalisée");
        const decoded = decodeJwt(authToken);
        console.log("👤 Utilisateur:", decoded?.preferred_username);
        console.log("📋 Scopes dans le token:", decoded?.scp);
        
        setAuthStatus("authenticated");
        
        // Utiliser le token directement pour Graph
        // Le token a les scopes Graph même si on demande la ressource personnalisée
        const graph = Client.init({
          authProvider: (done) => done(null, authToken),
        });
        
        setGraphClient(graph);
        const accounts = msalInstance.getAllAccounts();
        if (accounts && accounts.length > 0) {
          setAccount(accounts[0]);
          initGraphClient(accounts[0]);
        }
        setError(null);
        
      } catch (err) {
        console.error("❌ Erreur d'authentification:", err);
        setAuthStatus("error");
        
        if (err.message?.includes("Invalid resource") || err.message?.includes("650057")) {
          setError("Configuration Azure AD manquante: La ressource personnalisée n'est pas configurée dans Azure AD. Vérifiez 'Exposer une API'.");
        } else {
          setError("Erreur d'authentification: " + (err.message || JSON.stringify(err)));
        }
      }
    };

    initializeTeams();
  }, []);
  function openTeamsAuthDialog() {
    microsoftTeams.authentication.authenticate({
      url: window.location.origin + "/auth.html",  // la page d’auth dans un dialog
      width: 600,
      height: 600,
      successCallback: (resultToken) => {
        console.log("✅ Auth dialog OK, token reçu:", resultToken);
  
        const account = msalInstance.getAllAccounts()[0];
        setAccount(account);
        initGraphClient(account);
      },
      failureCallback: (reason) => {
        console.error("❌ Auth dialog erreur:", reason);
        setError(reason);
      }
    });
  }
  
  function initGraphClient(activeAccount) {
    const msalProvider = {
      getAccessToken: async () => {
        const request = { ...loginRequest, account: activeAccount };
        const response = await msalInstance.acquireTokenSilent(request).catch(async (e) => {
          return await msalInstance.acquireTokenPopup(request);
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

  /** ✅ Tester la connexion Graph */
  async function testGraphConnection() {
    if (!client) return;

    try {
      setLoading(true);
      // Tester avec une requête simple
      const user = await client.api('/me').get();
      console.log("✅ Test Graph réussi:", user.displayName);
      setError(null);
      return true;
    } catch (err) {
      console.error("❌ Test Graph échoué:", err);
      setError("Erreur Graph: " + (err.message || err));
      return false;
    } finally {
      setLoading(false);
    }
  }

  async function listPdfs() {
   // if (!graphClient) return;
  
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

  function closePreview() {
    setPreviewUrl(null);
  }

  return (
    <div style={{ padding: 20, fontFamily: "Segoe UI, sans-serif" }}>
      <h2>📄 MultiHealth — PDF Viewer</h2>
      
      <div style={{ marginBottom: 20, padding: 10, backgroundColor: "#f5f5f5", borderRadius: 4 }}>
        <p>
          <strong>Site:</strong> {siteUrl}<br />
          <strong>Dossier:</strong> {folderPath || "/ (racine)"}<br />
          <strong>Statut:</strong> {authStatus === "authenticated" ? "✅ Authentifié" : 
                                  authStatus === "teams_initialized" ? "🔄 Authentification..." : 
                                  authStatus === "error" ? "❌ Erreur" : "🔄 Initialisation..."}
        </p>
      </div>
      {!account && (
  <button
    onClick={openTeamsAuthDialog}
    style={{
      padding: "10px 20px",
      backgroundColor: "#0078d4",
      color: "white",
      border: "none",
      borderRadius: 4,
      cursor: "pointer",
      marginBottom: 20
    }}
  >
    🔐 Se connecter à Microsoft Graph
  </button>
)}
      <div style={{ marginBottom: 10 }}>
        <button 
          onClick={listPdfs} 
          disabled={!graphClient || loading}
          style={{
            padding: "10px 20px",
            backgroundColor: graphClient ? "#0078d4" : "#ccc",
            color: "white",
            border: "none",
            borderRadius: 4,
            cursor: graphClient ? "pointer" : "not-allowed",
            marginRight: 10
          }}
        >
          {loading ? "⏳ Chargement..." : "📂 Lister les PDF"}
        </button>

        {graphClient && (
          <button 
            onClick={testGraphConnection}
            disabled={loading}
            style={{
              padding: "10px 15px",
              backgroundColor: "#6c757d",
              color: "white",
              border: "none",
              borderRadius: 4,
              cursor: "pointer"
            }}
          >
            Test Graph
          </button>
        )}
      </div>

      {error && (
        <div style={{ 
          color: "red", 
          backgroundColor: "#ffe6e6",
          padding: 10,
          borderRadius: 4,
          marginTop: 10,
          border: "1px solid #ffcccc"
        }}>
          ❌ {error}
        </div>
      )}

      {!graphClient && !error && (
        <div style={{ 
          color: "#666", 
          padding: 10,
          marginTop: 10
        }}>
          🔄 {authStatus === "teams_initialized" ? 
              "Authentification avec ressource personnalisée..." : 
              "Initialisation de Teams..."}
        </div>
      )}

      {files.length > 0 && (
        <div style={{ marginTop: 20 }}>
          <h3>📋 Fichiers PDF ({files.length})</h3>
          <ul style={{ listStyle: "none", padding: 0 }}>
            {files.map(f => (
              <li key={f.id} style={{ 
                padding: "10px", 
                border: "1px solid #ddd", 
                marginBottom: 5,
                borderRadius: 4,
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center"
              }}>
                <span>📄 {f.name}</span>
                <button 
                  onClick={() => previewFile(f)}
                  disabled={loading}
                  style={{
                    padding: "5px 10px",
                    backgroundColor: "#28a745",
                    color: "white",
                    border: "none",
                    borderRadius: 3,
                    cursor: "pointer"
                  }}
                >
                  {loading ? "⏳" : "Aperçu"}
                </button>
              </li>
            ))}
          </ul>
        </div>
      )}

      {previewUrl && (
        <div style={{ marginTop: 20 }}>
          <div style={{ 
            display: "flex", 
            justifyContent: "space-between", 
            alignItems: "center",
            marginBottom: 10 
          }}>
            <h3>👁️ Aperçu PDF</h3>
            <button 
              onClick={closePreview}
              style={{
                padding: "5px 10px",
                backgroundColor: "#dc3545",
                color: "white",
                border: "none",
                borderRadius: 3,
                cursor: "pointer"
              }}
            >
              Fermer
            </button>
          </div>
          <iframe 
            src={previewUrl} 
            title="preview"
            style={{ 
              width: "100%", 
              height: "80vh", 
              border: "1px solid #ddd",
              borderRadius: 4
            }} 
          />
        </div>
      )}
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
export default App;