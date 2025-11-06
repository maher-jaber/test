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
        
        const authToken = await microsoftTeams.authentication.getAuthToken({
          resources: ["https://graph.microsoft.com"]
        });
        
        console.log("✅ Token obtenu");
        const decoded = decodeJwt(authToken);
        console.log("👤 Utilisateur:", decoded?.preferred_username);
        
        setAuthStatus("authenticated");
        
        const graph = Client.init({
          authProvider: (done) => done(null, authToken),
        });
        
        setGraphClient(graph);
        setError(null);
        
      } catch (err) {
        console.error("❌ Erreur d'authentification:", err);
        setAuthStatus("error");
        setError("Erreur d'authentification: " + (err.message || JSON.stringify(err)));
      }
    };

    initializeTeams();
  }, []);

  /** ✅ Lister les PDFs avec la méthode éprouvée */
  async function listPdfs() {
    if (!graphClient) {
      setError("Client Graph non initialisé");
      return;
    }
  
    setLoading(true);
    setError(null);
  
    try {
      console.log("🔍 Début de la recherche...");
      console.log("🔗 Site URL:", siteUrl);
      console.log("📁 Folder Path:", folderPath);

      // Extraire l'hostname et le chemin du site
      const hostname = new URL(siteUrl).hostname;
      const pathParts = new URL(siteUrl).pathname.split("/").filter(Boolean);
      const sitePath = pathParts.slice(1).join("/");

      console.log("🌐 Hostname:", hostname);
      console.log("🛣️ Site Path:", sitePath);

      // 1️⃣ Récupérer le site SharePoint
      const site = await graphClient.api(`/sites/${hostname}:/sites/${sitePath}`).get();
      console.log("✅ Site ID:", site.id);
      console.log("🏷️ Site Name:", site.displayName);

      // 2️⃣ Récupérer TOUTES les drives (bibliothèques documentaires)
      const drives = await graphClient.api(`/sites/${site.id}/drives`).get();
      console.log("📂 Drives trouvés:", drives.value.map(d => ({ name: d.name, id: d.id })));

      // 3️⃣ Trouver la drive qui contient les documents
      let driveId = null;
      let selectedDrive = null;
      
      for (let d of drives.value) {
        if (d.name.toLowerCase().includes("document") || d.driveType === "documentLibrary") {
          driveId = d.id;
          selectedDrive = d;
          console.log("✅ Drive sélectionnée:", d.name, d.id);
          break;
        }
      }

      // Fallback: prendre la première drive si aucune trouvée
      if (!driveId && drives.value.length > 0) {
        driveId = drives.value[0].id;
        selectedDrive = drives.value[0];
        console.log("🔄 Fallback sur la première drive:", selectedDrive.name);
      }

      if (!driveId) throw new Error("❌ Aucune bibliothèque de documents trouvée.");

      // 4️⃣ Construire le chemin API pour le dossier
      let apiPath;
      if (folderPath && folderPath.trim() !== "") {
        // Nettoyer le chemin du dossier
        let cleanFolderPath = folderPath.trim();
        if (!cleanFolderPath.startsWith('/')) {
          cleanFolderPath = '/' + cleanFolderPath;
        }
        apiPath = `/drives/${driveId}/root:${cleanFolderPath}:/children`;
      } else {
        apiPath = `/drives/${driveId}/root/children`;
      }

      console.log("🛣️ Chemin API Graph:", apiPath);

      // 5️⃣ Récupérer les fichiers
      const response = await graphClient.api(apiPath).get();
      console.log("📄 Éléments trouvés:", response.value.length);

      // 6️⃣ Filtrer les PDFs
      const pdfFiles = response.value.filter(f => {
        const isPdf = f.file && f.name.toLowerCase().endsWith(".pdf");
        if (isPdf) {
          console.log("📋 PDF trouvé:", f.name);
        }
        return isPpdf;
      });

      setFiles(pdfFiles);
      
      if (pdfFiles.length === 0) {
        setError("Aucun fichier PDF trouvé dans le dossier: " + (folderPath || "racine"));
      } else {
        console.log("✅ PDFs trouvés:", pdfFiles.length);
      }

    } catch (err) {
      console.error("❌ Erreur lors du listage:", err);
      
      // Gestion d'erreur détaillée
      if (err.statusCode === 404) {
        setError("Dossier non trouvé. Vérifiez le chemin: " + folderPath);
      } else if (err.statusCode === 403) {
        setError("Accès refusé. Vérifiez les permissions SharePoint.");
      } else if (err.message?.includes("Invalid hostname")) {
        setError("URL du site SharePoint invalide: " + siteUrl);
      } else {
        setError("Erreur: " + (err.message || JSON.stringify(err)));
      }
    } finally {
      setLoading(false);
    }
  }

  /** ✅ Aperçu PDF avec l'API Graph */
  async function previewFile(file) {
    if (!graphClient) {
      setError("Client Graph non initialisé");
      return;
    }

    setLoading(true);
    setError(null);

    try {
      console.log("👀 Génération de l'aperçu pour:", file.name);

      // Utiliser l'API de preview de Graph
      const previewResult = await graphClient
        .api(`/drives/${file.parentReference.driveId}/items/${file.id}/preview`)
        .post({});

      console.log("✅ Résultat preview:", previewResult);

      if (previewResult && previewResult.getUrl) {
        setPreviewUrl(previewResult.getUrl);
      } else {
        throw new Error("Impossible de générer l'aperçu");
      }

    } catch (err) {
      console.error("❌ Erreur preview:", err);
      setError("Impossible d'ouvrir le PDF: " + (err.message || JSON.stringify(err)));
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