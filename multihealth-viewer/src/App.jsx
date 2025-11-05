import React, { useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { Client } from "@microsoft/microsoft-graph-client";
import 'regenerator-runtime/runtime';
import * as microsoftTeams from "@microsoft/teams-js";

const AZURE_APP_ID = "1135fab5-62e8-4cb1-b472-880c477a8812";

function App() {
  const [graphClient, setGraphClient] = useState(null);
  const [files, setFiles] = useState([]);
  const [previewUrl, setPreviewUrl] = useState(null);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);

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
        
        // Essayer d'abord avec la ressource personnalisée
        try {
          const authToken = await microsoftTeams.authentication.getAuthToken({
            resources: [`api://test-rssn.onrender.com/1135fab5-62e8-4cb1-b472-880c477a8812`]
          });
          console.log("✅ Token avec ressource personnalisée obtenu");
          initializeGraphClient(authToken);
          
        } catch (customResourceError) {
          console.log("❌ Ressource personnalisée échouée, tentative avec Graph directement...");
          
          // Fallback: utiliser directement Microsoft Graph
          const authToken = await microsoftTeams.authentication.getAuthToken({
            resources: [`https://graph.microsoft.com`]
          });
          console.log("✅ Token Graph obtenu");
          initializeGraphClient(authToken);
        }
        
      } catch (err) {
        console.error("❌ Erreur d'authentification:", err);
        setError("Erreur d'authentification: " + (err.message || err));
      }
    };

    const initializeGraphClient = (token) => {
      const graph = Client.init({
        authProvider: (done) => done(null, token),
      });
      setGraphClient(graph);
      setError(null);
    };

    initializeTeams();
  }, []);

  /** ✅ Lister les PDF */
  async function listPdfs() {
    if (!graphClient) {
      setError("Client Graph non initialisé");
      return;
    }

    setLoading(true);
    setError(null);

    try {
      console.log("📂 Recherche du site...");
      
      const hostname = new URL(siteUrl).hostname;
      const site = await graphClient.api(`/sites/${hostname}:`).get();
      console.log("✅ Site trouvé:", site.displayName);

      const drives = await graphClient.api(`/sites/${site.id}/drives`).get();
      const drive = drives.value.find(d => 
        d.name.toLowerCase().includes("document")
      ) || drives.value[0];
      
      if (!drive) throw new Error("Aucune bibliothèque trouvée");

      const apiPath = folderPath ? 
        `/drives/${drive.id}/root:${folderPath}:/children` :
        `/drives/${drive.id}/root/children`;
      
      const response = await graphClient.api(apiPath).get();
      const pdfFiles = response.value.filter(f => f.file && f.name.toLowerCase().endsWith(".pdf"));
      
      setFiles(pdfFiles);
      if (pdfFiles.length === 0) setError("Aucun PDF trouvé");

    } catch (err) {
      console.error("❌ Erreur:", err);
      setError("Erreur: " + (err.message || "Impossible de charger les fichiers"));
    } finally {
      setLoading(false);
    }
  }

  /** ✅ Preview PDF */
  async function previewFile(file) {
    if (!graphClient) return;

    setLoading(true);
    setError(null);

    try {
      const preview = await graphClient
        .api(`/drives/${file.parentReference.driveId}/items/${file.id}/preview`)
        .post({ viewer: "web" });

      setPreviewUrl(preview.getUrl);
      
    } catch (err) {
      setError("Impossible de générer l'aperçu: " + (err.message || err));
    } finally {
      setLoading(false);
    }
  }

  return (
    <div style={{ padding: 20, fontFamily: "Segoe UI, sans-serif" }}>
      <h2>📄 MultiHealth — PDF Viewer</h2>
      
      <div style={{ marginBottom: 20 }}>
        <p><strong>Site:</strong> {siteUrl}<br />
        <strong>Dossier:</strong> {folderPath || "/ (racine)"}</p>
      </div>

      <button 
        onClick={listPdfs} 
        disabled={!graphClient || loading}
        style={{
          padding: "10px 20px",
          backgroundColor: graphClient ? "#0078d4" : "#ccc",
          color: "white",
          border: "none",
          borderRadius: 4,
          cursor: graphClient ? "pointer" : "not-allowed"
        }}
      >
        {loading ? "⏳ Chargement..." : "📂 Lister les fichiers PDF"}
      </button>

      {error && (
        <div style={{ color: "red", backgroundColor: "#ffe6e6", padding: 10, borderRadius: 4, marginTop: 10 }}>
          ❌ {error}
        </div>
      )}

      {files.length > 0 && (
        <div style={{ marginTop: 20 }}>
          <h3>📋 Fichiers PDF ({files.length})</h3>
          <ul style={{ listStyle: "none", padding: 0 }}>
            {files.map(f => (
              <li key={f.id} style={{ padding: "10px", border: "1px solid #ddd", marginBottom: 5, borderRadius: 4 }}>
                <span>📄 {f.name}</span>
                <button onClick={() => previewFile(f)} style={{ marginLeft: 10 }}>
                  Aperçu
                </button>
              </li>
            ))}
          </ul>
        </div>
      )}

      {previewUrl && (
        <div style={{ marginTop: 20 }}>
          <iframe 
            src={previewUrl} 
            title="preview"
            style={{ width: "100%", height: "80vh", border: "1px solid #ddd" }} 
          />
        </div>
      )}
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
export default App;