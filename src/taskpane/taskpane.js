/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 */

/* global document, Office */

// Import polyfills at the top
import "core-js/stable";
import "regenerator-runtime/runtime";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("get-properties-button").onclick = getItemProperties;
    document.getElementById("compose-reply-button").onclick = composeReply;
    
    // Load initial email info
    loadInitialData();
  }
});

/**
 * Load initial email data when taskpane opens
 */
function loadInitialData() {
  Office.context.mailbox.item.subject.getAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      document.getElementById("subject-text").textContent = result.value || "No subject";
    }
  });

  // Get sender information
  if (Office.context.mailbox.item.from) {
    Office.context.mailbox.item.from.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const from = result.value;
        document.getElementById("from-text").textContent = 
          from.displayName ? `${from.displayName} (${from.emailAddress})` : from.emailAddress;
      }
    });
  }

  // Get date
  if (Office.context.mailbox.item.dateTimeCreated) {
    const date = Office.context.mailbox.item.dateTimeCreated;
    document.getElementById("date-text").textContent = date.toLocaleDateString();
  }
}

/**
 * Get detailed email properties
 */
function getItemProperties() {
  const item = Office.context.mailbox.item;
  let properties = [];

  // Basic properties
  properties.push(`Item Type: ${item.itemType}`);
  properties.push(`Item ID: ${item.itemId || 'Not available'}`);
  
  // Get conversation ID
  if (item.conversationId) {
    properties.push(`Conversation ID: ${item.conversationId}`);
  }

  // Get categories
  if (item.categories) {
    item.categories.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const categories = result.value.length > 0 ? result.value.join(', ') : 'None';
        properties.push(`Categories: ${categories}`);
        updateResultArea(properties);
      }
    });
  } else {
    updateResultArea(properties);
  }

  // Get body preview
  if (item.body) {
    item.body.getAsync(Office.CoercionType.Text, { asyncContext: properties }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const bodyPreview = result.value.substring(0, 200) + '...';
        result.asyncContext.push(`Body Preview: ${bodyPreview}`);
        updateResultArea(result.asyncContext);
      }
    });
  }
}

/**
 * Compose a reply to the current email
 */
function composeReply() {
  Office.context.mailbox.item.displayReplyForm({
    htmlBody: "<p>This is an automated reply from the development add-in!</p><p>Best regards,<br/>Your Add-in</p>",
    attachments: []
  });
  
  updateResultArea(["Reply window opened!"]);
}

/**
 * Update the result display area
 * @param {string[]} properties - Array of properties to display
 */
function updateResultArea(properties) {
  const resultText = document.getElementById("result-text");
  resultText.innerHTML = properties.map(prop => `<p>${prop}</p>`).join('');
}

/**
 * Example function to get attachments (if any)
 */
function getAttachments() {
  const item = Office.context.mailbox.item;
  if (item.attachments && item.attachments.length > 0) {
    const attachmentInfo = item.attachments.map(att => 
      `${att.name} (${att.attachmentType})`
    ).join(', ');
    return `Attachments: ${attachmentInfo}`;
  }
  return "Attachments: None";
}

/**
 * Example function to work with custom properties
 */
function setCustomProperty() {
  Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const customProps = result.value;
      customProps.set("myCustomProperty", "Hello from add-in!");
      customProps.saveAsync((saveResult) => {
        if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
          updateResultArea(["Custom property saved successfully!"]);
        }
      });
    }
  });
}