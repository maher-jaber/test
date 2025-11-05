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

  /** ✅ Lister les PDF */
  /** ✅ Lister les PDF via SharePoint REST API */
async function listPdfs() {
  if (!graphClient) {
    setError("Client Graph non initialisé");
    return;
  }

  setLoading(true);
  setError(null);

  try {
    console.log("📂 Tentative via SharePoint REST API...");
    
    // Utiliser l'API SharePoint REST directement avec le token Graph
    const authToken = await microsoftTeams.authentication.getAuthToken();
    
    // Construire l'URL SharePoint REST
    const siteUri = new URL(siteUrl);
    const webUrl = `${siteUrl}/_api/web`;
    
    // Obtenir le dossier
    const folderRelativeUrl = folderPath || 'Shared Documents';
    const apiUrl = `${webUrl}/GetFolderByServerRelativeUrl('${folderRelativeUrl}')/Files`;
    
    console.log("🔍 URL SharePoint REST:", apiUrl);

    const response = await fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'Authorization': `Bearer ${authToken}`
      }
    });

    if (!response.ok) {
      throw new Error(`Erreur SharePoint: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    const allFiles = data.d.results;
    
    console.log("📄 Fichiers trouvés via SharePoint:", allFiles.length);

    // Filtrer les PDFs et formater comme l'ancienne structure
    const pdfFiles = allFiles.filter(f => f.Name.toLowerCase().endsWith('.pdf'));
    
    const formattedFiles = pdfFiles.map(f => ({
      id: f.UniqueId,
      name: f.Name,
      webUrl: f.ServerRelativeUrl,
      file: { mimeType: 'application/pdf' },
      parentReference: {
        driveId: 'sharepoint' // Valeur par défaut
      }
    }));

    setFiles(formattedFiles);
    
    if (pdfFiles.length === 0) {
      setError("Aucun fichier PDF trouvé dans ce dossier");
    } else {
      console.log("✅ PDFs trouvés via SharePoint:", pdfFiles.map(f => f.Name));
    }

  } catch (err) {
    console.error("❌ Erreur SharePoint REST:", err);
    
    // Fallback: Essayer avec l'ancienne méthode Graph
    console.log("🔄 Tentative de fallback avec Graph API...");
    await listPdfsWithGraphFallback();
  } finally {
    setLoading(false);
  }
}

/** ✅ Fallback avec Graph API */
async function listPdfsWithGraphFallback() {
  try {
    console.log("📂 Fallback: Recherche du site via Graph...");
    
    const hostname = new URL(siteUrl).hostname;
    const site = await graphClient.api(`/sites/${hostname}:`).get();
    console.log("✅ Site trouvé:", site.displayName);

    const drives = await graphClient.api(`/sites/${site.id}/drives`).get();
    const drive = drives.value.find(d => 
      d.name.toLowerCase().includes("document") || 
      d.name.toLowerCase().includes("documents") ||
      d.name.toLowerCase().includes("general")
    ) || drives.value[0];
    
    if (!drive) {
      throw new Error("Aucune bibliothèque de documents trouvée");
    }

    const apiPath = folderPath ? 
      `/drives/${drive.id}/root:${folderPath}:/children` :
      `/drives/${drive.id}/root/children`;
    
    const response = await graphClient.api(apiPath).get();
    const pdfFiles = response.value.filter(f => f.file && f.name.toLowerCase().endsWith(".pdf"));
    
    setFiles(pdfFiles);
    
    if (pdfFiles.length === 0) {
      setError("Aucun fichier PDF trouvé dans ce dossier");
    }

  } catch (err) {
    console.error("❌ Erreur fallback Graph:", err);
    setError("Impossible de charger les fichiers: " + (err.message || "Vérifiez l'URL et les permissions"));
  }
}

/** ✅ Preview PDF direct depuis SharePoint */
async function previewFile(file) {
  setLoading(true);
  setError(null);

  try {
    console.log("👀 Génération de l'aperçu direct pour:", file.name);
    
    // Méthode 1: URL directe vers le PDF dans SharePoint
    let pdfUrl;
    
    if (file.webUrl) {
      // Si on a l'URL SharePoint directe
      pdfUrl = file.webUrl.startsWith('http') ? file.webUrl : `${siteUrl}${file.webUrl}`;
    } else {
      // Fallback: Construire l'URL
      const encodedFileName = encodeURIComponent(file.name);
      const folderSegment = folderPath ? `${folderPath}/` : '';
      pdfUrl = `${siteUrl}/${folderSegment}${encodedFileName}`;
    }
    
    console.log("🔗 URL PDF directe:", pdfUrl);
    
    // S'assurer que c'est bien une URL absolue
    if (!pdfUrl.startsWith('http')) {
      pdfUrl = `${siteUrl}${pdfUrl.startsWith('/') ? '' : '/'}${pdfUrl}`;
    }
    
    setPreviewUrl(pdfUrl);
    console.log("✅ Aperçu direct configuré");

  } catch (err) {
    console.error("❌ Erreur preview direct:", err);
    
    // Fallback: Essayer avec l'ancienne méthode Graph
    console.log("🔄 Fallback: Aperçu via Graph API...");
    await previewFileWithGraphFallback(file);
  } finally {
    setLoading(false);
  }
}

/** ✅ Fallback preview avec Graph API */
async function previewFileWithGraphFallback(file) {
  try {
    console.log("👀 Fallback: Génération de l'aperçu Graph pour:", file.name);
    
    const preview = await graphClient
      .api(`/drives/${file.parentReference.driveId}/items/${file.id}/preview`)
      .post({
        viewer: "web",
        allowEdit: false,
        page: '1'
      });

    console.log("✅ URL d'aperçu Graph générée");
    setPreviewUrl(preview.getUrl);
    
  } catch (err) {
    console.error("❌ Erreur preview Graph fallback:", err);
    setError("Impossible de générer l'aperçu: " + (err.message || err));
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