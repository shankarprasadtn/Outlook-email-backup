# M365 Outlook & SharePoint Sync Utility

A high-performance, modern Windows desktop application designed to monitor your Outlook folders, extract email attachments with custom filters, dynamically rename files, and synchronize them directly to SharePoint Online or local OneDrive paths.

It also comes with a built-in **Live Web Sync Console** that runs completely in your web browser!

---

## 🚀 Key Features

* **Dual Scan Modes**:
  * **Outlook Desktop COM**: Scan your local running Outlook application (zero setup required).
  * **Exchange Online (MS Graph)**: Scan directly from Microsoft cloud servers without needing Outlook open.
* **Dual Upload Targets**:
  * **OneDrive Local Sync**: Fast file copy directly into your local OneDrive/SharePoint synced directory.
  * **SharePoint Cloud (Graph API)**: Upload files directly to the cloud library using OAuth2/MSAL.
* **Intelligent File Bifurcation**:
  * Normal files are stored in your primary directory (e.g. `2026`).
  * Agreements containing keywords like `spare` or `replacement` are automatically routed to the sibling folder `Spare 2026`.
* **Attachment Filtering**:
  * Prevents clutter by skipping email signature images, logos, and unrelated files unless they match your specified filters.
* **Background Scheduling**:
  * Set a sync polling interval (1 to 60 minutes) and minimize the utility to the Windows system tray.
* **Live Web Sync**:
  * Run the synchronization process directly from the web browser!

---

## 💻 Standalone Executable

You can download and run the precompiled standalone executable directly without installing Python:
* **Download**: Go to the [dist/](dist/) folder in this repository and download [OutlookSharePointSync.exe](dist/OutlookSharePointSync.exe).

---

## 🌐 Web App (GitHub Pages)

The landing page and Live Web Sync client is hosted at:
👉 **[https://shankarprasadtn.github.io/Outlook-email-backup/](https://shankarprasadtn.github.io/Outlook-email-backup/)**

---

## 🛠️ Configuration and Azure AD Setup

To connect to M365 services, you must register a Client ID (Application ID) and Tenant ID:

### 1. Azure Portal Registration
1. Visit the [Microsoft Entra ID Portal](https://entra.microsoft.com/) / [Azure Portal](https://portal.azure.com/).
2. Navigate to **Microsoft Entra ID** -> **App registrations** -> **New registration**.
3. Set your Redirect URIs:
   * **Single-page application (SPA)**: `https://shankarprasadtn.github.io/Outlook-email-backup/` (required for Web Sync)
   * **Public client/native**: `http://localhost` (required for Desktop Sync)
4. Click **Register** and copy your **Application (client) ID** and **Directory (tenant) ID**.

### 2. Graph API Permissions
Under **API Permissions**, select **Add a permission** -> **Microsoft Graph** -> **Delegated permissions**:
* `Mail.ReadWrite` (read inbox and mark processed emails as read)
* `Files.ReadWrite.All` (upload to SharePoint library)
* `Sites.ReadWrite.All` (look up SharePoint target site folders)
* Click **Grant admin consent** for your organization.

---

## 📦 Developer Installation (From Source)

1. Clone this repository:
   ```bash
   git clone https://github.com/shankarprasadtn/Outlook-email-backup.git
   cd Outlook-email-backup
   ```
2. Run the automated installer batch script:
   ```bash
   run.bat
   ```
   This will create a virtual environment (`.venv`), install dependencies, and launch the application.

3. Re-compile into a single executable if you make changes:
   ```bash
   compile.bat
   ```

---

## 📄 License

This project is licensed under the MIT License.
