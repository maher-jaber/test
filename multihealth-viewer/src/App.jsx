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
      console.log("📁 Dossier:", folderPath || "racine");
  
      // Tester d'abord la connexion Graph
      const testOk = await testGraphConnection();
      if (!testOk) {
        throw new Error("La connexion Graph a échoué");
      }
  
      // Méthode plus simple : utiliser search pour trouver les PDFs
      console.log("🔍 Recherche des PDFs via search...");
      
      // Construction de la requête de recherche
      const searchQuery = `site:${siteUrl} ${folderPath ? `path:${folderPath}` : ''} filetype:pdf`;
      
      console.log("🔎 Query de recherche:", searchQuery);
      
      const searchResult = await graphClient
        .api('/search/query')
        .version('beta')
        .post({
          requests: [
            {
              entityTypes: ['driveItem'],
              query: {
                queryString: searchQuery
              },
              fields: [
                'id',
                'name',
                'webUrl',
                'file',
                'parentReference',
                'size',
                'lastModifiedDateTime',
                '@microsoft.graph.downloadUrl'
              ]
            }
          ]
        });
  
      console.log("📊 Résultat search:", searchResult);
  
      if (searchResult.value && searchResult.value[0] && searchResult.value[0].hitsContainers) {
        const hits = searchResult.value[0].hitsContainers[0].hits;
        console.log("📄 Fichiers trouvés via search:", hits.length);
  
        const pdfFiles = hits.map(hit => hit.resource);
        setFiles(pdfFiles);
  
        if (pdfFiles.length === 0) {
          setError("Aucun fichier PDF trouvé dans ce dossier");
        } else {
          console.log("✅ PDFs trouvés:", pdfFiles.map(f => f.name));
        }
      } else {
        // Fallback : méthode directe avec l'URL du site
        console.log("🔄 Fallback: méthode directe...");
        await listPdfsDirectMethod();
      }
  
    } catch (err) {
      console.error("❌ Erreur recherche search:", err);
      
      // Fallback vers la méthode directe
      try {
        console.log("🔄 Tentative de fallback avec méthode directe...");
        await listPdfsDirectMethod();
      } catch (fallbackError) {
        console.error("❌ Erreur fallback:", fallbackError);
        setError("Impossible de charger les fichiers: " + (fallbackError.message || "Vérifiez l'URL et les permissions"));
      }
    } finally {
      setLoading(false);
    }
  }
  
  /** ✅ Méthode directe pour lister les PDFs */
  async function listPdfsDirectMethod() {
    try {
      console.log("🔍 Méthode directe: recherche du site...");
      
      const siteUri = new URL(siteUrl);
      const hostname = siteUri.hostname;
      
      console.log("🌐 Hostname:", hostname);
  
      // Obtenir le site root
      const site = await graphClient.api(`/sites/${hostname}:`).get();
      console.log("✅ Site root trouvé:", site.displayName, "- ID:", site.id);
  
      // Obtenir tous les sites pour trouver le bon
      const sites = await graphClient.api('/sites').get();
      console.log("🏢 Sites disponibles:", sites.value.map(s => ({ name: s.displayName, url: s.webUrl })));
  
      // Trouver le site qui correspond à notre URL
      const targetSite = sites.value.find(s => 
        s.webUrl && s.webUrl.toLowerCase().includes(hostname.toLowerCase())
      );
  
      if (!targetSite) {
        throw new Error(`Aucun site trouvé pour ${siteUrl}`);
      }
  
      console.log("🎯 Site cible trouvé:", targetSite.displayName, "- ID:", targetSite.id);
  
      // Maintenant utiliser le drive du site
      const drive = await graphClient.api(`/sites/${targetSite.id}/drive`).get();
      console.log("📁 Drive trouvé:", drive.name, "- ID:", drive.id);
  
      // Lister les fichiers
      const apiPath = folderPath ? 
        `/sites/${targetSite.id}/drive/root:${folderPath}:/children` :
        `/sites/${targetSite.id}/drive/root/children`;
      
      console.log("🛣️ Chemin API final:", apiPath);
      
      const response = await graphClient.api(apiPath).get();
      console.log("📄 Éléments bruts:", response.value);
  
      // Filtrer les PDF
      const pdfFiles = response.value.filter(f => {
        const isPdf = f.file && f.name.toLowerCase().endsWith(".pdf");
        if (isPdf) {
          console.log("📋 PDF trouvé:", f.name, "- Taille:", f.size, "- ID:", f.id);
        }
        return isPdf;
      });
  
      setFiles(pdfFiles);
      
      if (pdfFiles.length === 0) {
        setError("Aucun fichier PDF trouvé dans ce dossier. Vérifiez que le dossier existe et contient des PDFs.");
      } else {
        console.log("✅ PDFs trouvés:", pdfFiles.length);
      }
  
    } catch (err) {
      console.error("❌ Erreur méthode directe:", err);
      
      let errorMessage = "Erreur: " + (err.message || "Impossible de charger les fichiers");
      
      if (err.statusCode === 403) {
        errorMessage = "Accès refusé. Vérifiez que l'application a les permissions 'Sites.Read.All' dans Azure AD.";
      } else if (err.statusCode === 404) {
        errorMessage = "Site ou dossier non trouvé. Vérifiez que l'URL du site SharePoint est correcte.";
      } else if (err.statusCode === 401) {
        errorMessage = "Token invalide. Problème d'authentification.";
      } else if (err.code === "itemNotFound") {
        errorMessage = "Dossier non trouvé. Vérifiez le chemin du dossier.";
      }
      
      throw new Error(errorMessage);
    }
  }
  
  /** ✅ Preview PDF avec Graph API */
  async function previewFile(file) {
    if (!graphClient) return;
  
    setLoading(true);
    setError(null);
  
    try {
      console.log("👀 Génération de l'aperçu pour:", file.name);
      console.log("📋 Fichier info:", {
        id: file.id,
        driveId: file.parentReference?.driveId,
        hasDownloadUrl: !!file['@microsoft.graph.downloadUrl']
      });
  
      // Essayer d'abord l'URL de téléchargement direct
      if (file['@microsoft.graph.downloadUrl']) {
        console.log("✅ Utilisation de l'URL de téléchargement direct");
        setPreviewUrl(file['@microsoft.graph.downloadUrl']);
        return;
      }
  
      // Sinon utiliser l'API preview
      console.log("🔄 Utilisation de l'API preview...");
      
      const driveId = file.parentReference?.driveId;
      if (!driveId) {
        throw new Error("Drive ID non trouvé pour le fichier");
      }
  
      const preview = await graphClient
        .api(`/drives/${driveId}/items/${file.id}/preview`)
        .post({
          viewer: "web",
          allowEdit: false,
          page: '1'
        });
  
      console.log("✅ URL d'aperçu générée:", preview.getUrl);
      setPreviewUrl(preview.getUrl);
      
    } catch (err) {
      console.error("❌ Erreur preview:", err);
      
      // Dernier recours : essayer de construire l'URL manuellement
      try {
        console.log("🔄 Tentative avec URL manuelle...");
        const manualUrl = `${siteUrl}/${folderPath ? folderPath + '/' : ''}${file.name}`;
        console.log("🔗 URL manuelle:", manualUrl);
        setPreviewUrl(manualUrl);
      } catch (manualError) {
        setError("Impossible de générer l'aperçu: " + (err.message || err));
      }
    } finally {
      setLoading(false);
    }
  }
  
  /** ✅ Preview PDF avec Graph API */
  async function previewFile(file) {
    if (!graphClient) return;
  
    setLoading(true);
    setError(null);
  
    try {
      console.log("👀 Génération de l'aperçu pour:", file.name);
      
      // Utiliser l'URL de téléchargement direct
      const downloadUrl = file['@microsoft.graph.downloadUrl'];
      
      if (downloadUrl) {
        console.log("✅ Utilisation de l'URL de téléchargement direct");
        setPreviewUrl(downloadUrl);
      } else {
        // Fallback sur l'API preview
        const preview = await graphClient
          .api(`/drives/${file.parentReference.driveId}/items/${file.id}/preview`)
          .post({
            viewer: "web",
            allowEdit: false,
            page: '1'
          });
  
        console.log("✅ URL d'aperçu générée");
        setPreviewUrl(preview.getUrl);
      }
      
    } catch (err) {
      console.error("❌ Erreur preview:", err);
      setError("Impossible de générer l'aperçu: " + (err.message || err));
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