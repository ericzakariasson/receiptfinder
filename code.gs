// Constants
const OPENAI_API_URL = "https://api.openai.com/v1/chat/completions";
const THEME_COLOR = "#267BFB";
const PARENT_FOLDER_NAME = "receiptfinder";
const SPREADSHEET_NAME = "Found Receipts";

// Property keys
const PROCESSED_IDS_KEY = "processedMessageIds";
const LAST_PROCESSED_THREAD_INDEX_KEY = "lastProcessedThreadIndex";
const SPREADSHEET_ID_KEY = "receiptfinderSpreadsheetId";

const SPREADSHEET_HEADERS = [
  "Timestamp",
  "Date",
  "Status",
  "Drive Link",
  "Gmail Link",
  "Notes",
];

const BATCH_SIZE = 20;
const MAX_EXECUTION_TIME = 5.5 * 60 * 1000; // 5.5 minutes in milliseconds

// Installation and setup
function onInstall(e) {
  onOpen(e);
  setupIfNeeded();
}

function onOpen(e) {
  createTimeDrivenTrigger();
}

function setupIfNeeded() {
  const setupDone =
    PropertiesService.getUserProperties().getProperty("setupDone");
  if (!setupDone) {
    createDriveFolderAndSpreadsheet();
    PropertiesService.getUserProperties().setProperty("setupDone", "true");
  }
}

function createTimeDrivenTrigger() {
  deleteExistingTriggers_("processEmails");
  ScriptApp.newTrigger("processEmails")
    .timeBased()
    .everyHours(1) // Run every hour
    .create();
}

function deleteExistingTriggers_(functionName) {
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// Main processing function
function processEmails() {
  Logger.log("Starting processEmails function");

  handlePendingTasks();

  const processedIds = getProcessedMessageIds();
  const lastProcessedThreadIndex = getLastProcessedThreadIndex();

  const threads = GmailApp.search("", lastProcessedThreadIndex, BATCH_SIZE);
  if (threads.length === 0) {
    Logger.log("No more threads to process. Resetting index.");
    setLastProcessedThreadIndex(0);
    return;
  }

  const parentFolder = getOrCreateParentFolder();
  const metadataEntries = [];
  const fileCreationTasks = [];

  const startTime = new Date().getTime();

  for (let i = 0; i < threads.length; i++) {
    if (isTimeUp(startTime)) {
      saveProgress(
        lastProcessedThreadIndex + i,
        metadataEntries,
        fileCreationTasks
      );
      return;
    }

    processThread(
      threads[i],
      processedIds,
      parentFolder,
      metadataEntries,
      fileCreationTasks
    );
  }

  finalizeBatch(
    metadataEntries,
    fileCreationTasks,
    lastProcessedThreadIndex,
    threads.length
  );
}

function handlePendingTasks() {
  const pendingMetadataEntries =
    PropertiesService.getUserProperties().getProperty("pendingMetadataEntries");
  const pendingFileCreationTasks =
    PropertiesService.getUserProperties().getProperty(
      "pendingFileCreationTasks"
    );

  if (pendingMetadataEntries && pendingFileCreationTasks) {
    const metadataEntries = JSON.parse(pendingMetadataEntries);
    const fileCreationTasks = JSON.parse(pendingFileCreationTasks);

    const createdFiles = batchCreateFiles(fileCreationTasks);
    updateMetadataWithFileLinks(metadataEntries, createdFiles);
    if (metadataEntries.length > 0) {
      batchWriteMetadata(metadataEntries);
    }

    PropertiesService.getUserProperties().deleteProperty(
      "pendingMetadataEntries"
    );
    PropertiesService.getUserProperties().deleteProperty(
      "pendingFileCreationTasks"
    );
  }
}

function processThread(
  thread,
  processedIds,
  parentFolder,
  metadataEntries,
  fileCreationTasks
) {
  const messages = thread.getMessages();
  for (const emailMessage of messages) {
    const messageId = emailMessage.getId();
    if (processedIds.includes(messageId)) {
      Logger.log("Message already processed: %s", messageId);
      continue;
    }

    const result = checkIfReceipt(
      emailMessage.getPlainBody(),
      emailMessage.getAttachments()
    );
    if (result && result.is_receipt) {
      handleReceiptEmail(
        emailMessage,
        result,
        parentFolder,
        metadataEntries,
        fileCreationTasks
      );
    }

    recordProcessedMessageId(messageId);
  }
}

function handleReceiptEmail(
  emailMessage,
  result,
  parentFolder,
  metadataEntries,
  fileCreationTasks
) {
  const {
    attachments_sufficient,
    has_links_to_receipt,
    receipt_links,
    needs_manual_review,
  } = result;
  let status = needs_manual_review ? "Needs Manual Review" : "Processed";
  let notes = has_links_to_receipt
    ? `Contains links to receipt: ${receipt_links.join(", ")}\n`
    : "";

  if (attachments_sufficient) {
    fileCreationTasks.push(
      ...prepareAttachmentTasks(emailMessage, parentFolder)
    );
  } else if (!attachments_sufficient && !has_links_to_receipt) {
    fileCreationTasks.push(prepareEmailPDFTask(emailMessage, parentFolder));
  }

  metadataEntries.push(
    prepareMetadataEntry(emailMessage, [], notes.trim(), status)
  );
}

function finalizeBatch(
  metadataEntries,
  fileCreationTasks,
  lastProcessedThreadIndex,
  processedCount
) {
  const createdFiles = batchCreateFiles(fileCreationTasks);
  updateMetadataWithFileLinks(metadataEntries, createdFiles);
  if (metadataEntries.length > 0) {
    batchWriteMetadata(metadataEntries);
  }

  if (processedCount > 0) {
    const newLastIndex = lastProcessedThreadIndex + processedCount;
    setLastProcessedThreadIndex(newLastIndex);
    Logger.log("Updated last processed thread index to %s", newLastIndex);
  }
}

// Helper functions
function isTimeUp(startTime) {
  return new Date().getTime() - startTime > MAX_EXECUTION_TIME;
}

function saveProgress(newIndex, metadataEntries, fileCreationTasks) {
  setLastProcessedThreadIndex(newIndex);
  PropertiesService.getUserProperties().setProperty(
    "pendingMetadataEntries",
    JSON.stringify(metadataEntries)
  );
  PropertiesService.getUserProperties().setProperty(
    "pendingFileCreationTasks",
    JSON.stringify(fileCreationTasks)
  );
  scheduleNextRun();
}

function scheduleNextRun() {
  deleteExistingTriggers_("processEmails");
  ScriptApp.newTrigger("processEmails")
    .timeBased()
    .after(1 * 60 * 60 * 1000) // 1 hour in milliseconds
    .create();
}

function getLastProcessedThreadIndex() {
  const index = PropertiesService.getUserProperties().getProperty(
    LAST_PROCESSED_THREAD_INDEX_KEY
  );
  return index ? parseInt(index, 10) : 0;
}

function setLastProcessedThreadIndex(index) {
  PropertiesService.getUserProperties().setProperty(
    LAST_PROCESSED_THREAD_INDEX_KEY,
    index.toString()
  );
}

function getProcessedMessageIds() {
  const processedData =
    PropertiesService.getUserProperties().getProperty(PROCESSED_IDS_KEY);
  return processedData ? JSON.parse(processedData) : [];
}

function recordProcessedMessageId(messageId) {
  const processedIds = getProcessedMessageIds();
  processedIds.push(messageId);
  PropertiesService.getUserProperties().setProperty(
    PROCESSED_IDS_KEY,
    JSON.stringify(processedIds)
  );
}

// File and folder management
function createDriveFolderAndSpreadsheet() {
  const parentFolder = getOrCreateParentFolder();
  const spreadsheet = getOrCreateSpreadsheet(parentFolder);
  PropertiesService.getUserProperties().setProperty(
    SPREADSHEET_ID_KEY,
    spreadsheet.getId()
  );
}

function getOrCreateParentFolder() {
  const folders = DriveApp.getFoldersByName(PARENT_FOLDER_NAME);
  return folders.hasNext()
    ? folders.next()
    : DriveApp.createFolder(PARENT_FOLDER_NAME);
}

function getOrCreateSpreadsheet(parentFolder) {
  const files = parentFolder.getFilesByName(SPREADSHEET_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  } else {
    const spreadsheet = SpreadsheetApp.create(SPREADSHEET_NAME);
    const file = DriveApp.getFileById(spreadsheet.getId());
    parentFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
    spreadsheet.getActiveSheet().appendRow(SPREADSHEET_HEADERS);
    return spreadsheet;
  }
}

function getOrCreateReceiptSheet() {
  const spreadsheetId =
    PropertiesService.getUserProperties().getProperty(SPREADSHEET_ID_KEY);
  if (!spreadsheetId) {
    createDriveFolderAndSpreadsheet();
    return getOrCreateReceiptSheet();
  }
  try {
    return SpreadsheetApp.openById(spreadsheetId).getActiveSheet();
  } catch (error) {
    console.error("Error opening spreadsheet:", error);
    createDriveFolderAndSpreadsheet();
    return getOrCreateReceiptSheet();
  }
}

// OpenAI API interaction
function getOpenAIApiKey() {
  const apiKey =
    PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) {
    throw new Error(
      "OpenAI API key not found in script properties. Please set the OPENAI_API_KEY property."
    );
  }
  return apiKey;
}

function checkIfReceipt(emailBody, attachments) {
  Logger.log("Starting checkIfReceipt function");
  Logger.log("Checking if email is a receipt");
  Logger.log(`Email body length: ${emailBody.length}`);
  Logger.log(`Number of attachments: ${attachments.length}`);

  const attachmentInfo = attachments.map((att) => ({
    name: att.getName(),
    type: att.getContentType(),
  }));
  Logger.log(`Attachment info: ${JSON.stringify(attachmentInfo)}`);

  const payload = {
    model: "gpt-4o",
    messages: [
      {
        role: "system",
        content:
          "You are an AI assistant specialized in analyzing email messages to determine if they are receipts. Your task is to extract specific metadata relevant to the receipt for further processing. Focus on identifying explicit links to receipts or invoices, not just any link in the email.",
      },
      {
        role: "user",
        content: `Please analyze the following email content and attachment information to determine if it represents a receipt. Provide the required metadata in your response.

Email Content:
${emailBody}

Attachments:
${JSON.stringify(attachmentInfo, null, 2)}

In your analysis:
1. Determine if this is a receipt based on the email content and attachments.
2. Check if the attachments themselves are sufficient as receipt documentation.
3. Look for explicit links to receipts or invoices mentioned in the email body.
4. Decide if manual review is needed.

Provide your analysis in the required JSON format.`,
      },
    ],
    response_format: {
      type: "json_schema",
      json_schema: {
        name: "receipt_analysis",
        schema: {
          type: "object",
          properties: {
            is_receipt: { type: "boolean" },
            attachments_sufficient: { type: "boolean" },
            has_links_to_receipt: { type: "boolean" },
            receipt_links: { type: "array", items: { type: "string" } },
            needs_manual_review: { type: "boolean" },
            explanation: { type: "string" },
          },
          required: [
            "is_receipt",
            "attachments_sufficient",
            "has_links_to_receipt",
            "receipt_links",
            "needs_manual_review",
            "explanation",
          ],
          additionalProperties: false,
        },
        strict: true,
      },
    },
  };

  Logger.log("Preparing to call OpenAI API");
  const options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: `Bearer ${getOpenAIApiKey()}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    Logger.log("Calling OpenAI API");
    const response = UrlFetchApp.fetch(OPENAI_API_URL, options);
    Logger.log(`OpenAI API response status: ${response.getResponseCode()}`);
    const data = JSON.parse(response.getContentText());
    if (data.error) {
      throw new Error(`OpenAI API Error: ${data.error.message}`);
    }
    const result = JSON.parse(data.choices[0].message.content);
    Logger.log(`OpenAI API result: ${JSON.stringify(result)}`);
    return result;
  } catch (error) {
    console.error("Error calling OpenAI API or parsing response:", error);
    Logger.log(`Error in checkIfReceipt: ${error.message}`);
    return {
      is_receipt: false,
      attachments_sufficient: false,
      has_links_to_receipt: false,
      receipt_links: [],
      needs_manual_review: true,
      explanation: "Error occurred during AI analysis: " + error.message,
    };
  }
}

// UI functions
function onHomepage(e) {
  const cardBuilder = CardService.newCardBuilder()
    .setHeader(
      CardService.newCardHeader()
        .setTitle("Receipt Finder")
        .setSubtitle("Simplify your receipt management")
        .setImageUrl("https://example.com/logo.png")
    )
    .addSection(
      CardService.newCardSection().addWidget(
        CardService.newTextParagraph().setText(
          "<b>Welcome to Receipt Finder</b><br><br>" +
            "This add-on processes receipts from your Gmail and organizes them in Google Drive.<br>" +
            "Get started by clicking the button below."
        )
      )
    )
    .addSection(
      CardService.newCardSection().addWidget(
        CardService.newButtonSet()
          .addButton(
            CardService.newTextButton()
              .setText("▶ Run Receipt Finder")
              .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
              .setBackgroundColor(THEME_COLOR)
              .setOnClickAction(
                CardService.newAction().setFunctionName("onManualTrigger")
              )
          )
          .addButton(
            CardService.newTextButton()
              .setText("⟳ Reset All States")
              .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
              .setBackgroundColor(THEME_COLOR)
              .setOnClickAction(
                CardService.newAction().setFunctionName("onResetProcessedIds")
              )
          )
      )
    );

  return cardBuilder.build();
}

function onManualTrigger(e) {
  Logger.log("Manual trigger activated by user");
  processSingleThread();
  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText(
        "Single thread processing completed."
      )
    )
    .build();
}

function onResetProcessedIds(e) {
  Logger.log("Reset processed IDs activated by user");
  clearAllStates();
  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText("All states have been reset.")
    )
    .build();
}

function clearAllStates() {
  PropertiesService.getUserProperties().deleteAllProperties();
  const scriptProperties = PropertiesService.getScriptProperties();
  const allScriptProperties = scriptProperties.getProperties();
  for (const key in allScriptProperties) {
    if (key !== SPREADSHEET_ID_KEY && key !== "OPENAI_API_KEY") {
      scriptProperties.deleteProperty(key);
    }
  }
  Logger.log(
    "All stored states have been cleared, except for the spreadsheet ID and OpenAI API key"
  );
}

// Utility functions
function getGmailLink(emailMessage) {
  return `https://mail.google.com/mail/u/0?messageId=${emailMessage.getId()}`;
}

function prepareMetadataEntry(emailMessage, attachmentLinks, notes, status) {
  const timestamp = new Date();
  const date = emailMessage.getDate();
  const gmailLink = getGmailLink(emailMessage);
  const subject = emailMessage.getSubject() || "No Subject"; // Fallback if subject is empty

  return {
    emailMessageId: emailMessage.getId(),
    rowData: SPREADSHEET_HEADERS.map((header) => {
      switch (header.toLowerCase()) {
        case "timestamp":
          return timestamp;
        case "date":
          return date;
        case "status":
          return status;
        case "drive link":
          return ""; // Placeholder to be updated later
        case "gmail link":
          return `=HYPERLINK("${gmailLink}", "${subject.replace(/"/g, '""')}")`;
        case "notes":
          return notes;
        default:
          return ""; // Leave custom columns empty
      }
    }),
  };
}

function batchWriteMetadata(entries) {
  if (entries.length === 0) return;
  const sheet = getOrCreateReceiptSheet();
  const rowDataArray = entries.map((entry) => entry.rowData);
  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      rowDataArray.length,
      rowDataArray[0].length
    )
    .setValues(rowDataArray);
}

function prepareAttachmentTasks(emailMessage, parentFolder) {
  const tasks = [];
  const attachments = emailMessage.getAttachments();
  const monthFolderName = Utilities.formatDate(
    emailMessage.getDate(),
    Session.getScriptTimeZone(),
    "yyyy-MM"
  );

  for (const attachment of attachments) {
    tasks.push({
      type: "attachment",
      attachment: attachment,
      folderName: monthFolderName,
      parentFolder: parentFolder,
      emailMessageId: emailMessage.getId(),
    });
  }

  return tasks;
}

function prepareEmailPDFTask(emailMessage, parentFolder) {
  const emailSubject = emailMessage.getSubject().replace(/[^a-zA-Z0-9 ]/g, "");
  const emailDate = Utilities.formatDate(
    emailMessage.getDate(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
  const pdfName = `${emailDate} - ${emailSubject}.pdf`;
  const monthFolderName = Utilities.formatDate(
    emailMessage.getDate(),
    Session.getScriptTimeZone(),
    "yyyy-MM"
  );

  return {
    type: "pdf",
    content: emailMessage.getBody(),
    name: pdfName,
    folderName: monthFolderName,
    parentFolder: parentFolder,
    emailMessageId: emailMessage.getId(),
  };
}

function batchCreateFiles(tasks) {
  const createdFiles = [];
  const folderCache = {};

  for (const task of tasks) {
    const monthFolderName = task.folderName;
    let monthFolder = folderCache[monthFolderName];
    if (!monthFolder) {
      monthFolder = getOrCreateMonthFolder(task.parentFolder, monthFolderName);
      folderCache[monthFolderName] = monthFolder;
    }

    if (task.type === "attachment") {
      const attachmentName = task.attachment.getName();

      // Check for duplicate files
      const existingFiles = monthFolder.getFilesByName(attachmentName);
      if (existingFiles.hasNext()) {
        Logger.log("Duplicate attachment found, skipping: %s", attachmentName);
        continue; // Skip this attachment
      }

      const file = monthFolder.createFile(task.attachment);
      Logger.log("Saved attachment: %s", attachmentName);
      createdFiles.push({
        emailMessageId: task.emailMessageId,
        name: file.getName(),
        url: file.getUrl(),
      });
    } else if (task.type === "pdf") {
      const pdfName = task.name;

      // Check for duplicate files
      const existingFiles = monthFolder.getFilesByName(pdfName);
      if (existingFiles.hasNext()) {
        Logger.log("Duplicate PDF found, skipping: %s", pdfName);
        continue; // Skip this PDF
      }

      const blob = Utilities.newBlob(task.content, "text/html")
        .getAs("application/pdf")
        .setName(pdfName);
      const file = monthFolder.createFile(blob);
      Logger.log("Saved email as PDF: %s", pdfName);
      createdFiles.push({
        emailMessageId: task.emailMessageId,
        name: file.getName(),
        url: file.getUrl(),
      });
    }
  }

  return createdFiles;
}

function updateMetadataWithFileLinks(metadataEntries, createdFiles) {
  // Map emailMessageId to file URLs
  const fileLinksMap = {};
  for (const file of createdFiles) {
    if (!fileLinksMap[file.emailMessageId]) {
      fileLinksMap[file.emailMessageId] = [];
    }
    fileLinksMap[file.emailMessageId].push(file.url);
  }

  for (const entry of metadataEntries) {
    const emailMessageId = entry.emailMessageId;
    const driveLinks = fileLinksMap[emailMessageId] || [];
    const driveLink = driveLinks.join(", ");

    // Update the "Drive Link" field in entry
    const driveLinkIndex = SPREADSHEET_HEADERS.indexOf("Drive Link");
    if (driveLinkIndex !== -1) {
      entry.rowData[driveLinkIndex] = driveLink;
    }
  }
}

function getOrCreateMonthFolder(parentFolder, monthFolderName) {
  const folders = parentFolder.getFoldersByName(monthFolderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(monthFolderName);
  }
}

function processSingleThread() {
  Logger.log("Starting processSingleThread function");

  const processedIds = getProcessedMessageIds();
  const lastProcessedThreadIndex = getLastProcessedThreadIndex();

  const threads = GmailApp.search("", lastProcessedThreadIndex, 1);
  if (threads.length === 0) {
    Logger.log("No more threads to process. Resetting index.");
    setLastProcessedThreadIndex(0);
    return;
  }

  const parentFolder = getOrCreateParentFolder();
  const metadataEntries = [];
  const fileCreationTasks = [];

  processThread(
    threads[0],
    processedIds,
    parentFolder,
    metadataEntries,
    fileCreationTasks
  );

  finalizeBatch(
    metadataEntries,
    fileCreationTasks,
    lastProcessedThreadIndex,
    1
  );
}
