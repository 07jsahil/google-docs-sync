/**
 * ==============================================================================
 * 🚀 Zero-Touch Publishing Bridge: Google Docs to GitHub Actions (TEMPLATE)
 * ==============================================================================
 * Shareable open-source template for readers & creators.
 *
 * HOW TO INSTALL IN 3 STEPS:
 * 1. Open your Google Doc.
 * 2. In the top menu, go to: Extensions > Apps Script.
 * 3. Delete any default code, paste this entire file, customize the CONFIG
 *    block below, and click Save (💾).
 * 4. Refresh your Google Doc. You will see a new menu: "🚀 Website Admin".
 * ==============================================================================
 */

// --- ⚙️ USER CONFIGURATION ---
const CONFIG = {
  // 1. Your GitHub username or organization name
  GITHUB_OWNER: "YOUR_GITHUB_USERNAME", // e.g., "07jsahil"

  // 2. Your repository name where your website code lives
  GITHUB_REPO: "YOUR_REPOSITORY_NAME",   // e.g., "eskayengichem"

  // 3. Where in your repository this document should be saved as .docx
  FILE_PATH: "Products for website.docx",

  // 4. Target branch to commit to
  BRANCH: "main",

  // 5. GitHub Actions workflow dispatch event name
  EVENT_TYPE: "build_site",

  // 6. Name of the custom menu in Google Docs
  MENU_NAME: "🚀 Website Admin",

  // 7. Notification email(s) separated by commas (receives confirmation when site builds)
  // Leave empty ("") to automatically email the Google Doc owner.
  NOTIFY_EMAILS: "" // e.g., "owner@example.com, developer@example.com"
};

/**
 * Automatically creates the custom menu whenever the Google Doc is opened.
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu(CONFIG.MENU_NAME)
    .addItem('🚀 Update Website Now', 'syncAndBuildWebsite')
    .addSeparator()
    .addItem('⚙️ Configure GitHub Token', 'promptSetGitHubToken')
    .addToUi();
}

/**
 * Interactive prompt allowing non-technical users to set or update their
 * GitHub Personal Access Token directly from Google Docs without digging
 * into Apps Script Project Settings.
 */
function promptSetGitHubToken() {
  const ui = DocumentApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  const existingToken = scriptProperties.getProperty('GITHUB_PAT');
  
  const statusMsg = existingToken 
    ? "A GitHub Token is currently configured (••••" + existingToken.slice(-4) + ").\nPaste a new token to update it, or leave blank to cancel:"
    : "Enter your GitHub Personal Access Token (classic with 'repo' scope):";

  const response = ui.prompt('GitHub Token Configuration', statusMsg, ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const token = response.getResponseText().trim();
    if (token) {
      scriptProperties.setProperty('GITHUB_PAT', token);
      ui.alert('Success', 'GitHub Token saved securely in Script Properties!', ui.ButtonSet.OK);
    }
  }
}

/**
 * Main publishing function: Exports Doc -> Pushes to GitHub -> Triggers Build -> Sends Confirmation.
 */
function syncAndBuildWebsite() {
  const ui = DocumentApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // 1. Verify Configuration Placeholders
  if (CONFIG.GITHUB_OWNER === "YOUR_GITHUB_USERNAME" || CONFIG.GITHUB_REPO === "YOUR_REPOSITORY_NAME") {
    ui.alert(
      "Configuration Needed",
      "Please update the CONFIG block at the top of the script with your real GITHUB_OWNER and GITHUB_REPO names.",
      ui.ButtonSet.OK
    );
    return;
  }

  // 2. Retrieve GitHub Token
  let pat = scriptProperties.getProperty('GITHUB_PAT');
  if (!pat) {
    const prompt = ui.prompt(
      'GitHub Token Required',
      'Please enter your GitHub Personal Access Token (needs "repo" scope):',
      ui.ButtonSet.OK_CANCEL
    );
    if (prompt.getSelectedButton() === ui.Button.OK && prompt.getResponseText().trim()) {
      pat = prompt.getResponseText().trim();
      scriptProperties.setProperty('GITHUB_PAT', pat);
    } else {
      ui.alert('Cancelled', 'Website update cancelled: GitHub Token is required.', ui.ButtonSet.OK);
      return;
    }
  }

  const doc = DocumentApp.getActiveDocument();
  const docId = doc.getId();
  const docName = doc.getName();

  try {
    // 3. Export active Google Doc as a .docx file using Google Drive REST API
    const mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    const exportUrl = `https://www.googleapis.com/drive/v3/files/${docId}/export?mimeType=${encodeURIComponent(mimeType)}`;
    
    const driveResponse = UrlFetchApp.fetch(exportUrl, {
      method: "GET",
      headers: {
        "Authorization": "Bearer " + ScriptApp.getOAuthToken(),
        "Accept": "application/octet-stream"
      },
      muteHttpExceptions: true
    });

    if (driveResponse.getResponseCode() !== 200) {
      throw new Error(`Google Drive export failed (HTTP ${driveResponse.getResponseCode()}): ${driveResponse.getContentText()}`);
    }

    const docxBlob = driveResponse.getBlob();
    const base64Content = Utilities.base64Encode(docxBlob.getBytes());

    // 4. Check for existing file SHA on GitHub (required to update existing files)
    const contentsUrl = `https://api.github.com/repos/${CONFIG.GITHUB_OWNER}/${CONFIG.GITHUB_REPO}/contents/${encodeURIComponent(CONFIG.FILE_PATH)}?ref=${CONFIG.BRANCH}`;
    const getFileResponse = UrlFetchApp.fetch(contentsUrl, {
      method: "GET",
      headers: {
        "Authorization": `token ${pat}`,
        "Accept": "application/vnd.github.v3+json"
      },
      muteHttpExceptions: true
    });

    let fileSha = null;
    if (getFileResponse.getResponseCode() === 200) {
      const fileData = JSON.parse(getFileResponse.getContentText());
      fileSha = fileData.sha;
    } else if (getFileResponse.getResponseCode() === 401) {
      throw new Error("GitHub Authentication Failed (HTTP 401). Check if your GITHUB_PAT is valid and has 'repo' scope.");
    } else if (getFileResponse.getResponseCode() === 404) {
      fileSha = null;
    }

    // 5. Commit & Push the .docx file to GitHub
    const commitPayload = {
      message: `Content Update: Synchronize "${docName}" from Google Docs`,
      content: base64Content,
      branch: CONFIG.BRANCH
    };
    if (fileSha) {
      commitPayload.sha = fileSha;
    }

    const putUrl = `https://api.github.com/repos/${CONFIG.GITHUB_OWNER}/${CONFIG.GITHUB_REPO}/contents/${encodeURIComponent(CONFIG.FILE_PATH)}`;
    const putResponse = UrlFetchApp.fetch(putUrl, {
      method: "PUT",
      headers: {
        "Authorization": `token ${pat}`,
        "Content-Type": "application/json",
        "Accept": "application/vnd.github.v3+json"
      },
      payload: JSON.stringify(commitPayload),
      muteHttpExceptions: true
    });

    if (putResponse.getResponseCode() !== 200 && putResponse.getResponseCode() !== 201) {
      throw new Error(`GitHub file upload failed (HTTP ${putResponse.getResponseCode()}): ${putResponse.getContentText()}`);
    }

    // 6. Trigger GitHub Actions build workflow via repository_dispatch
    const dispatchUrl = `https://api.github.com/repos/${CONFIG.GITHUB_OWNER}/${CONFIG.GITHUB_REPO}/dispatches`;
    const dispatchPayload = {
      event_type: CONFIG.EVENT_TYPE,
      client_payload: {
        document_name: docName,
        triggered_at: new Date().toISOString()
      }
    };

    const dispatchResponse = UrlFetchApp.fetch(dispatchUrl, {
      method: "POST",
      headers: {
        "Authorization": `token ${pat}`,
        "Content-Type": "application/json",
        "Accept": "application/vnd.github.v3+json"
      },
      payload: JSON.stringify(dispatchPayload),
      muteHttpExceptions: true
    });

    if (dispatchResponse.getResponseCode() !== 204) {
      throw new Error(`GitHub Actions dispatch trigger failed (HTTP ${dispatchResponse.getResponseCode()}): ${dispatchResponse.getContentText()}`);
    }

    // 7. Send Email Notification
    const targetEmail = CONFIG.NOTIFY_EMAILS || Session.getEffectiveUser().getEmail();
    const repoUrl = `https://github.com/${CONFIG.GITHUB_OWNER}/${CONFIG.GITHUB_REPO}`;
    const emailSubject = `Website Update Triggered: "${docName}"`;
    const emailBody = `Hello,\n\nThe Google Doc "${docName}" was synchronized to GitHub, and your cloud build has been triggered successfully.\n\nSummary:\n- Repository: ${repoUrl}\n- File: ${CONFIG.FILE_PATH}\n- Target Branch: ${CONFIG.BRANCH}\n- Timestamp: ${new Date().toLocaleString()}\n\nYour changes should be live on your production website in ~30 seconds.\n\nBest regards,\nAutomated Zero-Touch Publishing Loop`;

    if (targetEmail) {
      MailApp.sendEmail(targetEmail, emailSubject, emailBody);
    }

    // 8. Positive UI Alert in Google Docs
    ui.alert(
      "🚀 Website Update Triggered!",
      `1. Document exported & pushed to GitHub (${CONFIG.FILE_PATH}).\n2. GitHub Actions cloud build started.\n3. Confirmation email sent to: ${targetEmail}\n\nYour live website will reflect these updates in ~30 seconds!`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    Logger.log(error.toString());
    ui.alert(
      "❌ Website Update Failed",
      error.message + "\n\nPlease check your settings and try again.",
      ui.ButtonSet.OK
    );
  }
}
