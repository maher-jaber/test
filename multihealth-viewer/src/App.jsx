import React, { useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import { Client } from "@microsoft/microsoft-graph-client";
import "regenerator-runtime/runtime";
import * as microsoftTeams from "@microsoft/teams-js";

function decodeJwt(token) {
  try {
    return JSON.parse(atob(token.split(".")[1]));
  } catch (e) {
    return null;
  }
}

function App() {
  const [client, setClient] = useState(null);
  const [previewUrl, setPreviewUrl] = useState(null);
  const [error, setError] = useState(null);
  const [authStatus, setAuthStatus] = useState("initializing");
  const [account, setAccount] = useState(null);
  const [loading, setLoading] = useState(false);

  const params = new URLSearchParams(window.location.search);
  const siteUrl = params.get("siteUrl") || "";
  const folderPath = params.get("folderPath") || "";

  /** ✅ Initialise Teams, récupère SSO Desktop sinon popup */
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
  }, []);

  /** ✅ Auth Microsoft (web) via popup */
  function openTeamsAuthDialog() {
    microsoftTeams.authentication.authenticate({
      url: window.location.origin + "/auth.html",
      width: 600,
      height: 600,
      successCallback: (accessToken) => {
        console.log("🔑 Token reçu via popup");
        initGraphClient(accessToken);

        setAccount({
          username: decodeJwt(accessToken)?.preferred_username,
          token: accessToken,
        });

        setAuthStatus("authenticated");
      },
      failureCallback: (reason) => {
        console.error("❌ Auth échouée:", reason);
        setError(reason);
      },
    });
  }

  /** ✅ Instancie Microsoft Graph */
  function initGraphClient(accessToken) {
    const graph = Client.init({
      authProvider: (done) => done(null, accessToken),
    });

    setClient(graph);
  }

  /** ✅ Lance automatiquement listPdfs quand client prêt */
  useEffect(() => {
    if (client) listPdfs();
  }, [client]);

  /** ✅ Récupère et ouvre automatiquement le premier PDF */
  async function listPdfs() {
    if (!client) return;

    setLoading(true);
    setError(null);

    try {
      const hostname = new URL(siteUrl).hostname;
      const pathParts = new URL(siteUrl).pathname.split("/").filter(Boolean);
      const sitePath = pathParts.slice(1).join("/");

      console.log("🔍 SITE:", hostname, sitePath);

      const site = await client.api(`/sites/${hostname}:/sites/${sitePath}`).get();

      const drives = await client.api(`/sites/${site.id}/drives`).get();

      const drive = drives.value.find((d) =>
        d.name.toLowerCase().includes("document")
      );

      if (!drive) throw new Error("❌ aucune bibliothèque Documents");

      const response = await client
        .api(`/drives/${drive.id}/root:${folderPath}:/children`)
        .get();

      const pdf = response.value.find(
        (f) => f.file && f.name.toLowerCase().endsWith(".pdf")
      );

      if (!pdf) throw new Error("❌ Aucun PDF trouvé dans ce dossier");

      console.log("📄 PDF détecté:", pdf.name);
      previewFile(pdf);
    } catch (e) {
      console.error(e);
      setError(e.message);
      setLoading(false);
    }
  }

  /** ✅ Récupère le lien d’aperçu */
  async function previewFile(item) {
    setError(null);
    setLoading(true);

    try {
      const res = await client
        .api(`/drives/${item.parentReference.driveId}/items/${item.id}/preview`)
        .post({});

      if (res?.getUrl) {
        setPreviewUrl(res.getUrl);
      } else {
        setError("Impossible de charger le PDF");
      }
    } catch (err) {
      setError(err.message);
    }

    setLoading(false);
  }

  return (
    <div
      style={{
        padding: 20,
        fontFamily: "'Segoe UI', sans-serif",
        background: "#f3f2f1",
        minHeight: "100vh",
      }}
    >
      <h2 style={{ marginBottom: 15 }}>📄 MultiHealth — PDF Viewer</h2>

      {!account && (
        <button
          onClick={openTeamsAuthDialog}
          style={{
            background: "#0078d4",
            padding: "10px 18px",
            borderRadius: 6,
            border: "none",
            color: "#fff",
            cursor: "pointer",
            fontWeight: 500,
          }}
        >
          🔐 Se connecter
        </button>
      )}

      {loading && <div style={{ marginTop: 10 }}>⏳ Chargement PDF…</div>}

    

      {previewUrl && (
        <iframe
          src={previewUrl}
          title="PDF Viewer"
          style={{
            width: "100%",
            height: "85vh",
            borderRadius: 8,
            border: "1px solid #ccc",
            marginTop: 14,
            background: "#fff",
          }}
        />
      )}
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
export default App;
