import React, { useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import { Client } from "@microsoft/microsoft-graph-client";
import * as microsoftTeams from "@microsoft/teams-js";
import * as msal from "@azure/msal-browser";
import "regenerator-runtime/runtime";

const AZURE_APP_ID = "1135fab5-62e8-4cb1-b472-880c477a8812";

const msalInstance = new msal.PublicClientApplication({
  auth: {
    clientId: AZURE_APP_ID,
  },
});

function decodeJwt(token) {
  try {
    return JSON.parse(atob(token.split(".")[1]));
  } catch {
    return null;
  }
}

function App() {
  const [graphClient, setGraphClient] = useState(null);
  const [files, setFiles] = useState([]);
  const [previewUrl, setPreviewUrl] = useState(null);
  const [error, setError] = useState(null);
  const [debugLogs, setDebugLogs] = useState([]);
  const [loading, setLoading] = useState(false);
  const [authStatus, setAuthStatus] = useState("initializing");

  const urlParams = new URLSearchParams(window.location.search);
  const siteUrl = urlParams.get("siteUrl") || "";
  const folderPath = urlParams.get("folderPath") || "";

  /** 🔧 helper debug */
  const log = (...msg) => {
    console.log(...msg);
    setDebugLogs((prev) => [...prev, msg.join(" ")]);
  };

  /** ✅ AUTHENTIFICATION SSO (Teams → JWT → MSAL → Token Graph) */
  useEffect(() => {
    const initTeamsSSO = async () => {
      try {
        log("🚀 Initialisation Teams...");
        await microsoftTeams.app.initialize();

        log("✅ Teams initialisé");
        setAuthStatus("teams_initialized");

        const teamsToken = await microsoftTeams.authentication.getAuthToken();
        const decoded = decodeJwt(teamsToken);
        log("👤 Utilisateur :", decoded?.preferred_username);

        const graphScopes = ["Files.Read", "Sites.Read.All", "User.Read"];

        log("🔐 Demande token Graph via MSAL...");

        const msalResult = await msalInstance.acquireTokenSilent({
          scopes: graphScopes,
          account: { username: decoded.preferred_username },
        });

        log("✅ Token Graph OK");

        const graph = Client.init({
          authProvider: (done) => done(null, msalResult.accessToken),
        });

        setGraphClient(graph);
        setAuthStatus("authenticated");
      } catch (err) {
        log("❌ Auth ERROR:", err);
        setAuthStatus("error");
        setError("Erreur d'authentification: " + (err.message || JSON.stringify(err)));
      }
    };

    initTeamsSSO();
  }, []);

  /** ✅ LISTE LES PDFs */
  async function listPdfs() {
    if (!graphClient) {
      setError("Client Graph non initialisé");
      return;
    }

    setLoading(true);
    setError(null);
    setFiles([]);

    try {
      log("📂 Début listage PDF");
      log("🌐 siteUrl:", siteUrl, " | folderPath:", folderPath);

      const hostname = new URL(siteUrl).hostname;
      const sitePath = new URL(siteUrl).pathname.split("/").filter(Boolean).slice(1).join("/");

      log("✅ Hostname:", hostname, " | sitePath:", sitePath);

      const site = await graphClient.api(`/sites/${hostname}:/sites/${sitePath}`).get();
      log("📌 site.id =", site.id);

      const drives = await graphClient.api(`/sites/${site.id}/drives`).get();
      log("📁 Drives trouvés:", JSON.stringify(drives.value.map(d => d.name)));

      let drive = drives.value.find(d => d.driveType === "documentLibrary") ?? drives.value[0];

      if (!drive) throw new Error("Aucune library trouvée");

      log("📌 Drive utilisée:", drive.name, "(", drive.id, ")");

      const cleanPath = folderPath ? `/root:/${folderPath}:/children` : `/root/children`;

      log("🛣️ API:", `/drives/${drive.id}${cleanPath}`);

      const response = await graphClient.api(`/drives/${drive.id}${cleanPath}`).get();

      const pdfFiles = response.value.filter(f => {
        const isPdf = f.file && f.name.toLowerCase().endsWith(".pdf");
        if (isPdf) log("➡️ PDF trouvé:", f.name);
        return isPdf;
      });

      setFiles(pdfFiles);

      if (pdfFiles.length === 0) setError("Aucun PDF trouvé 💡");

      log("✅ Fin listage:", pdfFiles.length, "PDFs trouvés");
    } catch (err) {
      log("❌ LIST ERROR:", err);
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }

  /** ✅ APERCU PDF */
  async function previewFile(file) {
    try {
      log("👀 Preview:", file.name);
      const preview = await graphClient.api(`/drives/${file.parentReference.driveId}/items/${file.id}/preview`).post({});
      setPreviewUrl(preview.getUrl);
    } catch (err) {
      log("❌ PREVIEW ERROR:", err);
      setError("Impossible d'ouvrir ce PDF");
    }
  }

  return (
    <div style={{ padding: 20, fontFamily: "Segoe UI" }}>
      <h2>📄 MultiHealth — PDF Viewer (debug)</h2>

      <p>
        <strong>Statut auth :</strong> {authStatus}
      </p>

      <button onClick={listPdfs} disabled={!graphClient || loading}>
        {loading ? "⏳ Chargement..." : "📂 Lister les PDF"}
      </button>

      {error && (
        <div style={{ background: "#ffdddd", padding: 10, marginTop: 10 }}>
          ❌ {error}
        </div>
      )}

      {files.length > 0 && (
        <ul>
          {files.map((f) => (
            <li key={f.id}>
              {f.name}
              <button onClick={() => previewFile(f)}>👁️ Aperçu</button>
            </li>
          ))}
        </ul>
      )}

      {previewUrl && (
        <iframe src={previewUrl} style={{ width: "100%", height: "70vh", marginTop: 20 }} />
      )}

      {/* ✅ OVERLAY DEBUG LOGS */}
      <div
        style={{
          position: "fixed",
          bottom: 0,
          right: 0,
          width: "370px",
          maxHeight: "250px",
          overflowY: "auto",
          background: "#111",
          color: "#0f0",
          fontSize: "12px",
          padding: "10px",
        }}
      >
        <strong>🟢 Debug logs :</strong>
        {debugLogs.map((l, i) => (
          <div key={i}>{l}</div>
        ))}
      </div>
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
export default App;
