import React, { useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { Client } from "@microsoft/microsoft-graph-client";
import 'regenerator-runtime/runtime';
import * as microsoftTeams from "@microsoft/teams-js";

const AZURE_APP_ID = "1135fab5-62e8-4cb1-b472-880c477a8812";


function decodeJwt(token) {
  try {
    return JSON.parse(atob(token.split('.')[1]));
  } catch (e) {
    return null;
  }
}

function App() {
  const [graphClient, setGraphClient] = useState(null);
  const [files, setFiles] = useState([]);
  const [previewUrl, setPreviewUrl] = useState(null);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);
  const [authStatus, setAuthStatus] = useState("initializing");

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

  /** ✅ Tester la connexion Graph */
  async function testGraphConnection() {
    if (!graphClient) return;

    try {
      setLoading(true);
      // Tester avec une requête simple
      const user = await graphClient.api('/me').get();
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

  /** ✅ Lister les PDFs via SharePoint REST API */
  async function listPdfs() {
    if (!siteUrl) {
      setError("URL du site manquante");
      return;
    }

    setLoading(true);
    setError(null);

    try {
      // Nettoyer le chemin du dossier
      const cleanFolderPath = folderPath.replace(/^\/+|\/+$/g, '');
      const relativePath = cleanFolderPath || 'Shared Documents';
      
      const apiUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${relativePath}')/Files`;
      
      console.log("🔍 Appel SharePoint:", apiUrl);

      const response = await fetch(apiUrl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
        },
        credentials: 'include'
      });

      if (!response.ok) {
        if (response.status === 403) {
          throw new Error("Accès refusé. Vérifiez vos permissions SharePoint.");
        } else if (response.status === 404) {
          throw new Error("Dossier non trouvé. Vérifiez le chemin.");
        }
        throw new Error(`Erreur ${response.status}: ${response.statusText}`);
      }

      const data = await response.json();
      const allFiles = data.d.results;
      
      console.log("📄 Fichiers bruts:", allFiles);

      // Filtrer les PDFs
      const pdfFiles = allFiles.filter(f => 
        f.Name.toLowerCase().endsWith('.pdf')
      );

      setFiles(pdfFiles.map(f => ({
        id: f.UniqueId,
        name: f.Name,
        webUrl: `${siteUrl}${f.ServerRelativeUrl}`,
        serverRelativeUrl: f.ServerRelativeUrl,
        lastModified: f.TimeLastModified,
        size: f.Length
      })));

      if (pdfFiles.length === 0) {
        setError("Aucun fichier PDF trouvé dans ce dossier");
      } else {
        console.log("✅ PDFs trouvés:", pdfFiles.length);
      }

    } catch (err) {
      console.error("❌ Erreur SharePoint:", err);
      setError(err.message || "Erreur lors du chargement des fichiers");
    } finally {
      setLoading(false);
    }
  }

  /** ✅ Preview PDF */
   /** ✅ Preview PDF direct depuis SharePoint */
   async function previewFile(file) {
    try {
      // URL directe vers le fichier dans SharePoint
      const pdfUrl = `${siteUrl}/${file.serverRelativeUrl}`;
      console.log("👀 Ouverture PDF:", pdfUrl);
      
      // Ouvrir dans un nouvel onglet ou intégrer
      setPreviewUrl(pdfUrl);
      
    } catch (err) {
      console.error("❌ Erreur preview:", err);
      setError("Impossible d'ouvrir le PDF: " + err.message);
    }
  }

  function closePreview() {
    setPreviewUrl(null);
  }

  return (
    <div style={{ padding: 20, fontFamily: "Segoe UI, sans-serif" }}>
      <h2>📄 MultiHealth — PDF Viewer (SharePoint Direct)</h2>
      
      <div style={{ marginBottom: 20, padding: 10, backgroundColor: "#f5f5f5", borderRadius: 4 }}>
        <p>
          <strong>Site:</strong> {siteUrl}<br />
          <strong>Dossier:</strong> {folderPath || "Shared Documents"}<br />
          <strong>Statut:</strong> {authStatus === "initialized" ? "✅ Prêt" : "🔄 Initialisation..."}
        </p>
      </div>

      <button 
        onClick={listPdfs} 
        disabled={loading || !siteUrl}
        style={{
          padding: "10px 20px",
          backgroundColor: siteUrl ? "#0078d4" : "#ccc",
          color: "white",
          border: "none",
          borderRadius: 4,
          cursor: siteUrl ? "pointer" : "not-allowed"
        }}
      >
        {loading ? "⏳ Chargement..." : "📂 Lister les PDF (SharePoint)"}
      </button>

      {/* Le reste du JSX reste identique */}
      {error && (
        <div style={{ color: "red", marginTop: 10 }}>
          ❌ {error}
        </div>
      )}

      {files.length > 0 && (
        <div style={{ marginTop: 20 }}>
          <h3>📋 Fichiers PDF ({files.length})</h3>
          <ul style={{ listStyle: "none", padding: 0 }}>
            {files.map(f => (
              <li key={f.id} style={{ padding: "10px", border: "1px solid #ddd", marginBottom: 5 }}>
                <span>📄 {f.name}</span>
                <button onClick={() => previewFile(f)}>
                  {loading ? "⏳" : "Aperçu"}
                </button>
              </li>
            ))}
          </ul>
        </div>
      )}

      {previewUrl && (
        <div style={{ marginTop: 20 }}>
          <button onClick={closePreview}>Fermer</button>
          <iframe 
            src={previewUrl} 
            style={{ width: "100%", height: "80vh", border: "1px solid #ddd" }} 
          />
        </div>
      )}
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
export default App;