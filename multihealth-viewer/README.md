
# Multi-Health-2 - Teams PDF Viewer (SharePoint)
This project is a Microsoft Teams Tab (React) that lists PDF files from a SharePoint folder and shows a preview using Microsoft Graph preview API.

## What is included
- React single-page app (simple, no UI framework)
- MSAL authentication (msal-browser)
- Uses Microsoft Graph to enumerate files & obtain preview URLs
- A sample Teams `manifest.json` + icons to create a Teams app package

## Quick start (developer)
1. `npm install`
2. Create an Azure AD App Registration with:
   - Redirect URI: `http://localhost:3000/`
   - Delegated API permissions: `Files.Read.All`, `Sites.Read.All`, `offline_access`, `openid`, `profile`
   - Grant admin consent for the tenant.
3. Copy `.env.example` to `.env` and fill `REACT_APP_CLIENT_ID` and `REACT_APP_TENANT_ID` and `REACT_APP_VALID_DOMAIN`.
4. `npm run start` to run locally (webpack dev server on :3000).

## Using inside Teams
- Host the built app on HTTPS and update `manifest.json` (in /teams-manifest) `contentUrl` and `validDomains` to point to your host.
- Zip `manifest.json`, `color.png`, `outline.png` and upload to Teams (Apps > Upload a custom app).

## How it works (short)
- The tab expects a query parameter `siteUrl` which is the SharePoint site URL (e.g. https://tenant.sharepoint.com/sites/YourSite)
- Optionally a `folderPath` query param, example: `?siteUrl=...&folderPath=/Shared Documents/PDFs`
- The app finds the drive and folder using Graph and lists files ending with `.pdf`.
- Clicking a file calls `POST /drives/{driveId}/items/{itemId}/preview` to get a preview URL and displays it in an iframe.

## Notes
- Replace the placeholder values in `/teams-manifest/manifest.json` before packaging for Teams.
- This code is a starter template — you can enhance UI, caching, error handling and pagination as needed.
