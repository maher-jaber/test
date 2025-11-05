import React, { useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { Client } from "@microsoft/microsoft-graph-client";
import 'regenerator-runtime/runtime';
import * as microsoftTeams from "@microsoft/teams-js";

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
        
        // Utiliser directement Microsoft Graph comme ressource
        const authToken = await microsoftTeams.authentication.getAuthToken({
          resources: ["https://graph.microsoft.com"]
        });
        
        console.log("✅ Token Microsoft Graph obtenu");
        const decoded = decodeJwt(authToken);
        console.log("👤 Utilisateur:", decoded?.preferred_username);
        console.log("🔑 Scopes:", decoded?.scp);
        
        setAuthStatus("authenticated");
        
        // Initialiser Graph client
        const graph = Client.init({
          authProvider: (done) => done(null, authToken),
        });
        
        setGraphClient(graph);
        setError(null);
        
      } catch (err) {
        console.error("❌ Erreur d'authentification:", err);
        setAuthStatus("error");
        setError("Erreur d'authentification: " + (err.message || JSON.stringify(err)));
        
        // Afficher plus de détails pour le débogage
        if (err.message?.includes("IncorrectConfiguration") || err.message?.includes("Access denied")) {
          setError(prev => prev + " - La ressource Microsoft Graph n'est pas autorisée. Vérifiez les permissions dans Azure AD.");
        }
      }
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
      
      // Extraire le hostname de l'URL du site
      const hostname = new URL(siteUrl).hostname;
      console.log("🔍 Hostname:", hostname);
      
      // Obtenir le site - utiliser l'approche avec ":" pour le hostname
      const site = await graphClient.api(`/sites/${hostname}:`).get();
      console.log("✅ Site trouvé:", site.displayName, "ID:", site.id);

      // Obtenir les drives (bibliothèques de documents)
      const drives = await graphClient.api(`/sites/${site.id}/drives`).get();
      console.log("📁 Drives disponibles:", drives.value.map(d => ({ name: d.name, id: d.id })));
      
      // Trouver le drive "Documents" ou le premier disponible
      const drive = drives.value.find(d => 
        d.name.toLowerCase().includes("document") || 
        d.name.toLowerCase().includes("documents") ||
        d.name.toLowerCase().includes("general")
      ) || drives.value[0];
      
      if (!drive) {
        throw new Error("Aucune bibliothèque de documents trouvée");
      }
      
      console.log("✅ Drive sélectionné:", drive.name, "ID:", drive.id);

      // Lister les fichiers dans le dossier spécifié
      let apiPath;
      if (folderPath && folderPath !== "/") {
        // Encoder le chemin du dossier pour l'URL
        const encodedPath = folderPath.startsWith("/") ? folderPath : `/${folderPath}`;
        apiPath = `/drives/${drive.id}/root:${encodedPath}:/children`;
      } else {
        apiPath = `/drives/${drive.id}/root/children`;
      }
      
      console.log("🔍 Chemin API:", apiPath);
      
      const response = await graphClient.api(apiPath).get();
      console.log("📄 Fichiers trouvés:", response.value.length);

      // Filtrer les PDF
      const pdfFiles = response.value.filter(f => f.file && f.name.toLowerCase().endsWith(".pdf"));
      setFiles(pdfFiles);
      
      if (pdfFiles.length === 0) {
        setError("Aucun fichier PDF trouvé dans ce dossier. Formats supportés: .pdf");
      } else {
        console.log("✅ PDFs trouvés:", pdfFiles.map(f => f.name));
      }

    } catch (err) {
      console.error("❌ Erreur lors de la liste des PDF:", err);
      let errorMessage = "Erreur: " + (err.message || "Impossible de charger les fichiers");
      
      // Messages d'erreur plus spécifiques
      if (err.statusCode === 403) {
        errorMessage = "Accès refusé. Vérifiez les permissions SharePoint.";
      } else if (err.statusCode === 404) {
        errorMessage = "Site ou dossier non trouvé. Vérifiez l'URL.";
      } else if (err.message?.includes("Invalid hostname")) {
        errorMessage = "URL du site invalide. Format attendu: https://votredomaine.sharepoint.com/sites/votresite";
      }
      
      setError(errorMessage);
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
      console.log("👀 Génération de l'aperçu pour:", file.name);
      
      const preview = await graphClient
        .api(`/drives/${file.parentReference.driveId}/items/${file.id}/preview`)
        .post({
          viewer: "web",
          allowEdit: false,
          page: '1'
        });

      console.log("✅ URL d'aperçu générée:", preview.getUrl);
      setPreviewUrl(preview.getUrl);
      
    } catch (err) {
      console.error("❌ Erreur preview:", err);
      setError("Impossible de générer l'aperçu: " + (err.message || err));
    } finally {
      setLoading(false);
    }
  }

  /** ✅ Fermer l'aperçu */
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
        {loading ? "⏳ Chargement..." : "📂 Lister les fichiers PDF"}
      </button>

      {graphClient && (
        <button 
          onClick={() => console.log("Client Graph:", graphClient)}
          style={{
            padding: "10px 20px",
            backgroundColor: "#6c757d",
            color: "white",
            border: "none",
            borderRadius: 4,
            cursor: "pointer"
          }}
        >
          Debug
        </button>
      )}

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
              "Authentification avec Microsoft Graph..." : 
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