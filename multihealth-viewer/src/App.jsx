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

  async function listPdfs() {
    if (!graphClient) {
      setError("Client Graph non initialisé");
      return;
    }
  
    setLoading(true);
    setError(null);
  
    try {
      console.log("📂 Début de la recherche...");
      console.log("🔗 Site URL:", siteUrl);
  
      // TEST 1: D'abord vérifier que le token Graph fonctionne
      try {
        const user = await graphClient.api('/me').get();
        console.log("✅ Test Graph /me réussi:", user.displayName);
      } catch (testError) {
        console.error("❌ Test Graph /me échoué:", testError);
        throw new Error("Token Graph invalide: " + testError.message);
      }
  
      // METHODE 1: Utiliser l'API SharePoint REST avec un token SharePoint
      console.log("🔄 Tentative avec SharePoint REST API...");
      
      // Obtenir un token spécifique pour ce site SharePoint
      const sharePointToken = await microsoftTeams.authentication.getAuthToken({
        resources: [siteUrl] // Utiliser l'URL du site comme ressource
      });
      
      console.log("✅ Token SharePoint obtenu");
  
      // Nettoyer le chemin du dossier
      let cleanFolderPath = folderPath || 'Shared Documents';
      if (cleanFolderPath.startsWith('/')) {
        cleanFolderPath = cleanFolderPath.substring(1);
      }
  
      // Construire l'URL SharePoint REST
      const apiUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${cleanFolderPath}')/Files`;
      
      console.log("🔍 URL SharePoint REST:", apiUrl);
  
      const response = await fetch(apiUrl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'Authorization': `Bearer ${sharePointToken}`
        }
      });
  
      if (!response.ok) {
        const errorText = await response.text();
        console.error("❌ Erreur SharePoint REST:", response.status, errorText);
        
        if (response.status === 403) {
          throw new Error("Accès refusé à SharePoint. Vérifiez les permissions.");
        } else if (response.status === 404) {
          throw new Error("Dossier non trouvé: " + cleanFolderPath);
        }
        throw new Error(`Erreur SharePoint ${response.status}: ${response.statusText}`);
      }
  
      const data = await response.json();
      const allFiles = data.d.results;
      
      console.log("📄 Fichiers trouvés via SharePoint REST:", allFiles.length);
  
      // Filtrer les PDFs et formater pour l'application
      const pdfFiles = allFiles.filter(f => {
        const isPdf = f.Name.toLowerCase().endsWith('.pdf');
        if (isPdf) {
          console.log("📋 PDF trouvé:", f.Name);
        }
        return isPdf;
      });
      
      // Formater les fichiers pour garder la compatibilité
      const formattedFiles = pdfFiles.map(f => ({
        id: f.UniqueId || f.Id,
        name: f.Name,
        webUrl: f.ServerRelativeUrl,
        file: { 
          mimeType: 'application/pdf' 
        },
        parentReference: {
          driveId: 'sharepoint'
        },
        '@microsoft.graph.downloadUrl': `${siteUrl}${f.ServerRelativeUrl}`,
        // Ajouter les propriétés SharePoint pour le preview
        _sharePointData: f
      }));
  
      setFiles(formattedFiles);
      
      if (pdfFiles.length === 0) {
        setError("Aucun fichier PDF trouvé dans le dossier: " + cleanFolderPath);
      } else {
        console.log("✅ PDFs trouvés:", pdfFiles.length);
      }
  
    } catch (err) {
      console.error("❌ Erreur principale:", err);
      
      // Fallback: Essayer avec Graph API si SharePoint REST échoue
      if (err.message.includes("Accès refusé") || err.message.includes("404")) {
        console.log("🔄 Fallback: tentative avec Graph API...");
        try {
          await listPdfsWithGraphFallback();
        } catch (graphError) {
          console.error("❌ Fallback Graph échoué:", graphError);
          setError(err.message + " | Fallback Graph: " + graphError.message);
        }
      } else {
        setError(err.message);
      }
    } finally {
      setLoading(false);
    }
  }
  
  /** ✅ Fallback avec Graph API */
  async function listPdfsWithGraphFallback() {
    console.log("🔍 Fallback Graph: recherche du site...");
    
    const siteUri = new URL(siteUrl);
    const hostname = siteUri.hostname;
    
    // Essayer de trouver le site par son hostname
    const site = await graphClient.api(`/sites/${hostname}:`).get();
    console.log("✅ Site trouvé via Graph:", site.displayName);
  
    // Utiliser le drive du site
    const drive = await graphClient.api(`/sites/${site.id}/drive`).get();
    console.log("📁 Drive trouvé:", drive.name);
  
    // Lister les fichiers
    const apiPath = folderPath ? 
      `/sites/${site.id}/drive/root:${folderPath}:/children` :
      `/sites/${site.id}/drive/root/children`;
    
    console.log("🛣️ Chemin API Graph:", apiPath);
    
    const response = await graphClient.api(apiPath).get();
    console.log("📄 Éléments Graph trouvés:", response.value.length);
  
    const pdfFiles = response.value.filter(f => f.file && f.name.toLowerCase().endsWith(".pdf"));
    setFiles(pdfFiles);
    
    if (pdfFiles.length === 0) {
      throw new Error("Aucun PDF trouvé avec Graph API");
    }
  }
  
  /** ✅ Preview PDF avec URL directe SharePoint */
  async function previewFile(file) {
    setLoading(true);
    setError(null);
  
    try {
      console.log("👀 Génération de l'aperçu pour:", file.name);
  
      // METHODE 1: URL directe SharePoint avec token
      let pdfUrl;
      
      if (file.webUrl) {
        // Si on a l'URL relative SharePoint
        pdfUrl = file.webUrl.startsWith('http') ? file.webUrl : `${siteUrl}${file.webUrl}`;
      } else if (file['@microsoft.graph.downloadUrl']) {
        // Si on a l'URL de téléchargement Graph
        pdfUrl = file['@microsoft.graph.downloadUrl'];
      } else {
        // Construire l'URL manuellement
        const encodedFileName = encodeURIComponent(file.name);
        const folderSegment = folderPath ? `${folderPath}/` : '';
        pdfUrl = `${siteUrl}/${folderSegment}${encodedFileName}`;
      }
  
      console.log("🔗 URL PDF:", pdfUrl);
  
      // Obtenir un token frais pour SharePoint
      const sharePointToken = await microsoftTeams.authentication.getAuthToken({
        resources: [siteUrl]
      });
  
      // Créer une URL avec le token pour l'authentification
      const previewUrlWithAuth = `${pdfUrl}?web=1`;
      
      console.log("✅ URL d'aperçu générée");
      setPreviewUrl(previewUrlWithAuth);
  
    } catch (err) {
      console.error("❌ Erreur preview:", err);
      setError("Impossible d'ouvrir le PDF: " + err.message);
    } finally {
      setLoading(false);
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