// ==========================================
// 1. SETTINGS
// ==========================================
const TRACKER_SHEET_NAME = "(PUBLIC) Lawsuits challenging AV laws"; 
const CHANGELOG_SHEET_NAME = "(Private) _Changelog";
const SUBSCRIBERS_SHEET_NAME = "(Private) Subscribers";

// State lookup helper 
const STATES = [
  "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", "Connecticut",
  "Delaware", "Florida", "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas",
  "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota",
  "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey",
  "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon",
  "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah",
  "Vermont", "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming", 
  "District of Columbia", "Puerto Rico", "Guam", "U.S. Virgin Islands", "Northern Mariana Islands"
];

// ==========================================
// 2. DYNAMIC HEADER FINDER
// ==========================================
function getTrackerConfig(sheet) {
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < Math.min(data.length, 15); i++) {
    const rowStrings = data[i].map(c => String(c).trim().toUpperCase());
    const caseIdx = rowStrings.indexOf("CASE ACTION");
    if (caseIdx !== -1) {
      return {
        headerRowIdx: i, 
        caseColIdx: caseIdx, 
        originalHeaders: data[i].map(c => String(c).trim()),
        data: data
      };
    }
  }
  return null;
}

// ==========================================
// 3. THE CUSTOM MENU
// ==========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📣 Tracker Alerts')
      .addItem('Send Pending Updates', 'sendDigest')
      .addToUi();
}

// ==========================================
// 4. THE SILENT LOGGER (Runs automatically)
// ==========================================
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== TRACKER_SHEET_NAME) return;

  const conf = getTrackerConfig(sheet);
  if (!conf) return;

  const rowIdx = e.range.getRow() - 1; 
  const colIdx = e.range.getColumn() - 1; 

  if (rowIdx <= conf.headerRowIdx) return;

  const caseName = String(conf.data[rowIdx][conf.caseColIdx]).trim();
  const colName = conf.originalHeaders[colIdx];

  if (!colName || !caseName) return;

  const oldVal = ('oldValue' in e) ? e.oldValue : "";
  const newVal = ('value' in e) ? e.value : "";

  // Generate a unique fingerprint for this specific row
  const upperHeaders = conf.originalHeaders.map(h => String(h).toUpperCase());
  const docketIdx = upperHeaders.indexOf("DISTRICT DOCKET");
  const dateIdx = upperHeaders.indexOf("DATE FILED");
  
  const docket = docketIdx !== -1 ? String(conf.data[rowIdx][docketIdx]).trim() : "";
  const dateF = dateIdx !== -1 ? String(conf.data[rowIdx][dateIdx]).trim() : "";
  const uniqueRowKey = caseName + "|||" + docket + "|||" + dateF;

  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHANGELOG_SHEET_NAME);
  if (logSheet) {
    logSheet.appendRow([new Date(), caseName, colName, oldVal, newVal, uniqueRowKey]);
  }
}

// ==========================================
// 5. WORD DIFF ENGINE
// ==========================================
function diffWords(oldStr, newStr, isHtml) {
  let prepO = String(oldStr || "").replace(/\n/g, " ¶ ");
  let prepN = String(newStr || "").replace(/\n/g, " ¶ ");
  let o = prepO.split(/\s+/).filter(Boolean);
  let n = prepN.split(/\s+/).filter(Boolean);
  
  let dp = Array(o.length + 1).fill(0).map(() => Array(n.length + 1).fill(0));
  for (let i = 1; i <= o.length; i++) {
    for (let j = 1; j <= n.length; j++) {
      if (o[i-1] === n[j-1]) dp[i][j] = dp[i-1][j-1] + 1;
      else dp[i][j] = Math.max(dp[i-1][j], dp[i][j-1]);
    }
  }
  
  let i = o.length, j = n.length;
  let res = [];
  while (i > 0 || j > 0) {
    if (i > 0 && j > 0 && o[i-1] === n[j-1]) {
      res.unshift({t: o[i-1], type: 0});
      i--; j--;
    } else if (j > 0 && (i === 0 || dp[i][j-1] >= dp[i-1][j])) {
      res.unshift({t: n[j-1], type: 1});
      j--;
    } else if (i > 0 && (j === 0 || dp[i][j-1] < dp[i-1][j])) {
      res.unshift({t: o[i-1], type: -1});
      i--;
    }
  }
  
  let out = [];
  for (let k = 0; k < res.length; k++) {
    let word = res[k].t;
    if (word === "¶") {
      out.push(isHtml ? "<br>" : "\n");
    } else {
      if (res[k].type === 0) {
        out.push(word);
      } else if (res[k].type === -1) {
        out.push(isHtml ? `<del style="color: #c026d3; text-decoration: line-through;">${word}</del>` : `~${word}~`);
      } else {
        // Changed to bold green text without underline
        out.push(isHtml ? `<ins style="font-weight:bold; color: #16a34a; text-decoration: none;">${word}</ins>` : `*${word}*`);
      }
    }
  }
  return out.join(" ").replace(/ <br> /g, "<br>").replace(/ \n /g, "\n").replace(/<br> /g, "<br>").replace(/\n /g, "\n");
}

// ==========================================
// 6. THE DIGEST SENDER
// ==========================================
function sendDigest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(CHANGELOG_SHEET_NAME);
  const subSheet = ss.getSheetByName(SUBSCRIBERS_SHEET_NAME);
  const trackerSheet = ss.getSheetByName(TRACKER_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();

  if (!logSheet) {
    ui.alert(`Error: Could not find the tab named "${CHANGELOG_SHEET_NAME}".`);
    return;
  }

  const lastRow = logSheet.getLastRow();
  if (lastRow === 0) {
    ui.alert("No pending updates to send!");
    return;
  }
  
  const conf = getTrackerConfig(trackerSheet);
  if (!conf) {
    ui.alert("Error: Could not find 'CASE ACTION' header in your tracker sheet.");
    return;
  }

  const lastCol = Math.max(logSheet.getLastColumn(), 5);
  const changes = logSheet.getRange(1, 1, lastRow, lastCol).getValues();
  
  const trackerData = conf.data;
  const trackerRichData = trackerSheet.getDataRange().getRichTextValues();
  const headers = conf.originalHeaders;
  const upperHeaders = headers.map(h => h.toUpperCase());

  // Group changes by Unique Row Key
  const changesByKey = {};
  changes.forEach(change => {
    let caseName = String(change[1]).trim();
    let column = String(change[2]).trim();
    let oldV = change[3];
    let newV = change[4];
    let uniqueKey = change[5] ? String(change[5]).trim() : caseName;
    
    if (!caseName) return; 
    if (!changesByKey[uniqueKey]) changesByKey[uniqueKey] = { caseName: caseName, changes: {} };
    
    if (changesByKey[uniqueKey].changes[column]) {
      changesByKey[uniqueKey].changes[column].newV = newV; 
    } else {
      changesByKey[uniqueKey].changes[column] = { oldV, newV };
    }
  });

  // Updated Intro string for email with legend
  let emailHtml = `
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; max-width: 900px; margin: 0 auto; color: #111827;">
      <p style="color: #6b7280; margin-bottom: 32px; font-size: 16px;">
        The following lawsuits have been updated in <a href="https://brennanschaffner.github.io/SMAVIC/" style="color: #2563eb; text-decoration: underline;">the tracker</a>. Changes are highlighted (<ins style="font-weight:bold; color: #16a34a; text-decoration: none;">additions</ins> <del style="color: #c026d3; text-decoration: line-through;">deletions</del>).
      </p>
  `;
  
  let slackBlocks = [
    {
      "type": "section",
      "text": { "type": "mrkdwn", "text": "The following lawsuits have been updated in <https://brennanschaffner.github.io/SMAVIC/|the tracker>." }
    },
    { "type": "divider" }
  ];

  // Data helpers
  const getValPlain = (rIdx, colName) => {
    let cIdx = upperHeaders.indexOf(colName.toUpperCase());
    return cIdx !== -1 ? trackerData[rIdx][cIdx] : "";
  };

  const getValRichHtml = (rIdx, colName) => {
    let cIdx = upperHeaders.indexOf(colName.toUpperCase());
    if (cIdx === -1) return "";
    let rtv = trackerRichData[rIdx][cIdx];
    if (!rtv) return "";
    let html = "";
    rtv.getRuns().forEach(run => {
      let t = run.getText().replace(/\n/g, "<br>");
      let u = run.getLinkUrl();
      if (u) html += `<a href="${u}" style="color: #2563eb; text-decoration: none;">${t}</a>`;
      else html += t;
    });
    return html || "-";
  };

  const getValRichSlack = (rIdx, colName) => {
    let cIdx = upperHeaders.indexOf(colName.toUpperCase());
    if (cIdx === -1) return "";
    let rtv = trackerRichData[rIdx][cIdx];
    if (!rtv) return "";
    let txt = "";
    rtv.getRuns().forEach(run => {
      let t = run.getText();
      let u = run.getLinkUrl();
      if (u) txt += `<${u}|${t}>`;
      else txt += t;
    });
    return txt || "-";
  };

  const getState = (statute) => {
    if(!statute) return "";
    let s = String(statute).toLowerCase();
    for (let i = 0; i < STATES.length; i++) {
      if (s.indexOf(STATES[i].toLowerCase()) === 0) return STATES[i];
    }
    return "";
  };

  const getHtmlField = (caseChanges, colName, rIdx) => {
    if (caseChanges[colName]) {
      let o = caseChanges[colName].oldV;
      let n = caseChanges[colName].newV;
      return diffWords(o, n, true) || "<em style='color:#9ca3af;'>[Cleared]</em>";
    }
    return getValRichHtml(rIdx, colName);
  };

  const getSlackField = (caseChanges, colName, rIdx) => {
    if (caseChanges[colName]) {
      let o = caseChanges[colName].oldV;
      let n = caseChanges[colName].newV;
      return diffWords(o, n, false) || "_[Cleared]_";
    }
    return getValRichSlack(rIdx, colName);
  };

  let casesProcessed = 0;
  const docketIdx = upperHeaders.indexOf("DISTRICT DOCKET");
  const dateIdx = upperHeaders.indexOf("DATE FILED");

  // Build Layouts for Each Unique Row
  for (const uniqueKey in changesByKey) {
    const caseData = changesByKey[uniqueKey];
    const caseName = caseData.caseName;
    const caseChanges = caseData.changes;

    let rIdx = -1;
    for (let i = 0; i < trackerData.length; i++) {
      let cName = String(trackerData[i][conf.caseColIdx]).trim();
      let dckt = docketIdx !== -1 ? String(trackerData[i][docketIdx]).trim() : "";
      let dt = dateIdx !== -1 ? String(trackerData[i][dateIdx]).trim() : "";
      
      let currentRowKey = cName + "|||" + dckt + "|||" + dt;
      
      if (currentRowKey === uniqueKey || (uniqueKey === cName && cName === caseName)) {
        rIdx = i;
        break;
      }
    }
    if (rIdx === -1) continue;
    casesProcessed++;

    // Base Fields
    let rawStatute = getValPlain(rIdx, "CHALLENGED STATUTE");
    let stateName = getState(rawStatute);
    let dateObj = getValPlain(rIdx, "DATE FILED");
    let dateStr = "-";
    if (dateObj instanceof Date) {
       dateStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MMM d, yyyy");
    } else if (dateObj) {
       dateStr = String(dateObj); 
    }

    // ---------- 1. BUILD EMAIL HTML ----------
    let statuteHtml = getHtmlField(caseChanges, "CHALLENGED STATUTE", rIdx);
    let claimsHtml = getHtmlField(caseChanges, "COMPLAINT", rIdx);
    let docketDisplayHtml = getHtmlField(caseChanges, "DISTRICT DOCKET", rIdx);
    let statusBoxHtml = getHtmlField(caseChanges, "LONG STATUS", rIdx);

    emailHtml += `
      <div style="background: #ffffff; border: 1px solid #e5e7eb; border-radius: 6px; padding: 24px; margin-bottom: 32px;">
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
          <tr>
            <td style="vertical-align: top;">
              <strong style="font-size: 18px; display: block;">${caseName}</strong>
            </td>
            <td style="vertical-align: top; text-align: right;">
              <span style="font-size: 14px; color: #6b7280; white-space: nowrap;">Filed: ${dateStr}</span>
            </td>
          </tr>
        </table>
    `;

    if (statusBoxHtml && statusBoxHtml !== "-") {
      emailHtml += `
        <div style="background: #f1f5f9; border-left: 4px solid #3b82f6; padding: 14px 18px; margin-bottom: 24px; font-size: 16px; font-weight: 500; color: #0f172a; line-height: 1.5;">
          ${statusBoxHtml}
        </div>
      `;
    }

    const gridFields = [
      ["Statute", "CHALLENGED STATUTE"],
      ["Claims", "COMPLAINT"],
      ["District", "DISTRICT"],
      ["District Docket", "DISTRICT DOCKET"],
      ["District Decision", "DISTRICT DECISION"],
      ["Appealing Party", "APPEALING PARTY"],
      ["Circuit", "CIRCUIT"],
      ["Appeals Docket", "APPEALS DOCKET(S)"],
      ["Circuit Decision", "CIRCUIT DECISION"],
      ["SCOTUS Docket", "SCOTUS DOCKET"],
      ["SCOTUS Opinion", "SCOTUS OPINION"],
      ["Notes", "NOTES"]
    ];

    emailHtml += `<table style="width: 100%; border-collapse: collapse; font-size: 14px;">`;
    gridFields.forEach(f => {
      let val = getHtmlField(caseChanges, f[1], rIdx);
      if (val !== "-" && val !== "") {
        // No explicit [Updated] tag for email since colors make it obvious
        emailHtml += `
          <tr>
            <td style="padding: 10px 16px 10px 0; width: 160px; font-weight: 600; font-size: 12px; color: #6b7280; text-transform: uppercase; vertical-align: top; border-bottom: 1px solid #f3f4f6;">${f[0]}</td>
            <td style="padding: 10px 0; vertical-align: top; border-bottom: 1px solid #f3f4f6;">${val}</td>
          </tr>
        `;
      }
    });
    emailHtml += `</table></div>`;

    // ---------- 2. BUILD SLACK BLOCKS ----------
    let caseBlocks = [];
    let caseText = `*${caseName}*\n\n`;

    // Always include Status 
    let statusBoxSlack = getSlackField(caseChanges, "LONG STATUS", rIdx);
    if (statusBoxSlack && statusBoxSlack !== "-") {
      let statusLabel = caseChanges["LONG STATUS"] ? `*[Changed]* *Status:*` : `*Status:*`;
      caseText += `${statusLabel}\n> ${statusBoxSlack.replace(/\n/g, "\n> ")}\n\n`;
    }

    // Always include Statute
    let statuteSlack = getSlackField(caseChanges, "CHALLENGED STATUTE", rIdx);
    if (statuteSlack && statuteSlack !== "-") {
      let statuteLabel = caseChanges["CHALLENGED STATUTE"] ? `*[Changed]* *Statute:*` : `*Statute:*`;
      caseText += `${statuteLabel}\n${statuteSlack}\n\n`;
    }

    // Include other fields ONLY if they changed
    gridFields.forEach(f => {
      if (f[1] === "CHALLENGED STATUTE") return; // Handled above
      if (caseChanges[f[1]]) {
        let val = getSlackField(caseChanges, f[1], rIdx);
        if (val && val !== "-") {
           caseText += `*[Changed]* *${f[0]}:*\n${val}\n\n`;
        }
      }
    });
    
    // Chunk just in case the length exceeds Slack's 3000-character limit per block
    if (caseText.length > 2800) {
        let chunks = caseText.match(/[\s\S]{1,2800}/g);
        chunks.forEach(c => {
            caseBlocks.push({ "type": "section", "text": { "type": "mrkdwn", "text": c.trim() }});
        });
    } else {
        caseBlocks.push({ "type": "section", "text": { "type": "mrkdwn", "text": caseText.trim() }});
    }
    
    caseBlocks.push({ "type": "divider" });
    slackBlocks = slackBlocks.concat(caseBlocks);
  }

  emailHtml += `</div>`;

  if (casesProcessed === 0) {
    ui.alert("Warning", "Changes were found in the log, but no matching cases exist in the tracker anymore. The log will be cleared.", ui.ButtonSet.OK);
    logSheet.clear();
    return;
  }

  // Sort out Emails and Webhooks
  let emails = [];
  let webhooks = [];
  
  if (subSheet && subSheet.getLastRow() > 0) {
    const rawList = subSheet.getRange(1, 1, subSheet.getLastRow(), 1).getValues().flat().filter(String);
    rawList.forEach(item => {
      let val = item.trim();
      if (val.indexOf("https://hooks.slack.com/") === 0) {
        webhooks.push(val);
      } else if (val.indexOf("@") !== -1) {
        emails.push(val);
      }
    });
  }

  let errors = [];

  // Send to Slack
  if (webhooks.length > 0) {
    const chunkSize = 40; 
    for (let i = 0; i < slackBlocks.length; i += chunkSize) {
      let chunk = slackBlocks.slice(i, i + chunkSize);
      let payload = { "blocks": chunk };
      let options = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(payload)
      };
      webhooks.forEach(hook => {
        try { UrlFetchApp.fetch(hook, options); } 
        catch(e) { errors.push("Slack Webhook failed: " + e.message); }
      });
    }
  }

  // Send Emails
  if (emails.length > 0) {
    try {
      MailApp.sendEmail({
        to: Session.getActiveUser().getEmail(), 
        bcc: emails.join(","), 
        subject: "Social Media Age Verification Litigation Update",
        htmlBody: emailHtml
      });
    } catch(e) {
      errors.push("Email failed: " + e.message);
    }
  }

  if (errors.length > 0) {
    ui.alert("Partial Failure", "Some messages could not be sent:\n\n" + errors.join("\n\n"), ui.ButtonSet.OK);
    return; 
  }

  logSheet.clear();
  ui.alert("Success!", "Updates have been perfectly formatted and sent to Slack and Email.", ui.ButtonSet.OK);
}