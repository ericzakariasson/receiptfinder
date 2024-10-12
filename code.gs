// Constants for OpenAI API
const OPENAI_API_URL = "https://api.openai.com/v1/chat/completions";
const OPENAI_API_KEY =
  "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"; // Replace with your OpenAI API key

// Label name for receipts
const RECEIPT_LABEL = "receipt";

// Name of the parent folder in Google Drive
const PARENT_FOLDER_NAME = "receiptfinder";

// Property key for storing processed message IDs
const PROCESSED_IDS_KEY = "processedMessageIds";

function onInstall(e) {
  createTrigger();
}

function createTrigger() {
  // Delete existing triggers to prevent duplicates
  deleteExistingTriggers_("processEmails");

  // Create a new time-based trigger
  ScriptApp.newTrigger("processEmails").timeBased().everyHours(1).create();
}

function deleteExistingTriggers_(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

function processEmails() {
  console.log("Starting processEmails function");

  const processedIds = getProcessedMessageIds();
  console.log("Processed Message IDs:", processedIds);

  const threads = GmailApp.search("is:unread");
  console.log("Found unread threads:", threads.length);

  // Get or create the parent folder
  const parentFolder = getOrCreateParentFolder();
  console.log("Using parent folder:", parentFolder.getName());

  for (const thread of threads) {
    const messages = thread.getMessages();
    console.log(
      "Processing thread ID:",
      thread.getId(),
      "with messages:",
      messages.length
    );

    for (const emailMessage of messages) {
      const messageId = emailMessage.getId();
      console.log("Processing message ID:", messageId);

      // Skip if the message has already been processed
      if (processedIds.includes(messageId)) {
        console.log("Message already processed:", messageId);
        continue;
      }

      if (!emailMessage.isUnread()) {
        console.log("Message is not unread, skipping:", messageId);
        continue;
      }

      const emailBody = emailMessage.getPlainBody();
      const isReceipt = checkIfReceipt(emailBody);
      console.log("Is message a receipt?", isReceipt);

      if (isReceipt) {
        // labelThread(thread); // Label the entire thread
        // console.log("Labeled thread as receipt:", thread.getId());

        saveAttachments(emailMessage, parentFolder);
        console.log("Saved attachments for message:", messageId);
      }

      // Mark the email as read
      // emailMessage.markRead();
      // console.log("Marked message as read:", messageId);

      // Record the message ID as processed
      recordProcessedMessageId(messageId);
      console.log("Recorded processed message ID:", messageId);
    }
  }
  console.log("Finished processing emails");
}

function checkIfReceipt(emailBody) {
  console.log("Checking if email is a receipt");

  const payload = {
    model: "gpt-4o-2024-08-06",
    messages: [
      {
        role: "system",
        content:
          "You are an AI assistant that analyzes emails to determine if they are receipts.",
      },
      {
        role: "user",
        content: `Please analyze the following email and determine if it is a receipt:\n\n${emailBody}`,
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
            explanation: { type: "string" },
          },
          required: ["is_receipt", "explanation"],
          additionalProperties: false,
        },
        strict: true,
      },
    },
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${OPENAI_API_KEY}`,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(OPENAI_API_URL, options);
    const data = JSON.parse(response.getContentText());
    const result = JSON.parse(data.choices[0].message.content);
    console.log("OpenAI API response:", result);

    return result.is_receipt;
  } catch (error) {
    console.error("Error calling OpenAI API or parsing response:", error);
    return false;
  }
}

function labelThread(thread) {
  let label = GmailApp.getUserLabelByName(RECEIPT_LABEL);
  if (!label) {
    label = GmailApp.createLabel(RECEIPT_LABEL);
  }
  thread.addLabel(label);
}

function saveAttachments(emailMessage, parentFolder) {
  const attachments = emailMessage.getAttachments();
  console.log("Found attachments:", attachments.length);

  if (attachments.length === 0) {
    console.log("No attachments to save for message:", emailMessage.getId());
    return;
  }

  const date = emailMessage.getDate();
  const monthFolderName = Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyy-MM"
  );
  console.log("Month folder name:", monthFolderName);

  // Get or create the month folder within the parent folder
  let monthFolder;
  const folders = parentFolder.getFoldersByName(monthFolderName);
  if (folders.hasNext()) {
    monthFolder = folders.next();
    console.log("Using existing month folder:", monthFolderName);
  } else {
    monthFolder = parentFolder.createFolder(monthFolderName);
    console.log("Created new month folder:", monthFolderName);
  }

  for (const attachment of attachments) {
    const attachmentName = attachment.getName();

    // Check if a file with the same name exists in the month folder
    const existingFiles = monthFolder.getFilesByName(attachmentName);
    if (existingFiles.hasNext()) {
      console.log("Duplicate attachment found, skipping:", attachmentName);
      continue; // Skip this attachment
    }

    monthFolder.createFile(attachment);
    console.log("Saved attachment:", attachmentName);
  }
}

function getOrCreateParentFolder() {
  console.log("Getting or creating parent folder:", PARENT_FOLDER_NAME);

  const folders = DriveApp.getFoldersByName(PARENT_FOLDER_NAME);
  if (folders.hasNext()) {
    console.log("Parent folder exists:", PARENT_FOLDER_NAME);
    return folders.next();
  } else {
    console.log("Creating parent folder:", PARENT_FOLDER_NAME);
    return DriveApp.createFolder(PARENT_FOLDER_NAME);
  }
}

function getProcessedMessageIds() {
  const properties = PropertiesService.getUserProperties();
  const processedData = properties.getProperty(PROCESSED_IDS_KEY);

  if (processedData) {
    return JSON.parse(processedData);
  } else {
    return [];
  }
}

function recordProcessedMessageId(messageId) {
  console.log("Recording processed message ID:", messageId);

  const properties = PropertiesService.getUserProperties();
  let processedData = properties.getProperty(PROCESSED_IDS_KEY);
  let processedIds;

  if (processedData) {
    processedIds = JSON.parse(processedData);
  } else {
    processedIds = [];
  }

  processedIds.push(messageId);

  // Update the property with the new list
  properties.setProperty(PROCESSED_IDS_KEY, JSON.stringify(processedIds));
  console.log("Updated processed message IDs");
}

function onHomepage(e) {
  // Create a card with a button to manually trigger the processEmails function
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("Gmail Receipt Finder"))
    .addSection(
      CardService.newCardSection()
        .addWidget(
          CardService.newTextParagraph().setText(
            "This add-on processes receipts from your Gmail and organizes them in Google Drive."
          )
        )
        .addWidget(
          CardService.newTextButton()
            .setText("Run Receipt Finder")
            .setOnClickAction(
              CardService.newAction().setFunctionName("onManualTrigger")
            )
        )
    )
    .build();
}

function onManualTrigger(e) {
  console.log("Manual trigger activated by user");
  processEmails();

  // Return a notification to the user
  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText("Receipt processing completed.")
    )
    .build();
}
