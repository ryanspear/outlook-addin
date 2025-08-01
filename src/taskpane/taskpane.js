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
  try {
    const item = Office.context.mailbox.item;
    
    // Get subject - this is a direct property, not async
    if (item.subject) {
      document.getElementById("subject-text").textContent = item.subject || "No subject";
    }

    // Get sender information - this IS async
    if (item.from) {
      item.from.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const from = result.value;
          document.getElementById("from-text").textContent = 
            from.displayName ? `${from.displayName} (${from.emailAddress})` : from.emailAddress;
        } else {
          document.getElementById("from-text").textContent = "Unable to load sender info";
        }
      });
    } else {
      // Fallback for compose mode or when from is not available
      document.getElementById("from-text").textContent = "Not available in compose mode";
    }

    // Get date - this is also a direct property
    if (item.dateTimeCreated) {
      const date = item.dateTimeCreated;
      document.getElementById("date-text").textContent = date.toLocaleDateString();
    } else {
      document.getElementById("date-text").textContent = "Date not available";
    }
    
  } catch (error) {
    console.error("Error loading initial data:", error);
    document.getElementById("subject-text").textContent = "Error loading data";
    document.getElementById("from-text").textContent = "Error loading data";
    document.getElementById("date-text").textContent = "Error loading data";
  }
}

/**
 * Get detailed email properties
 */
function getItemProperties() {
  try {
    const item = Office.context.mailbox.item;
    let properties = [];

    // Basic properties - these are direct properties
    properties.push(`Item Type: ${item.itemType}`);
    properties.push(`Item ID: ${item.itemId || 'Not available'}`);
    properties.push(`Subject: ${item.subject || 'No subject'}`);
    
    // Get conversation ID
    if (item.conversationId) {
      properties.push(`Conversation ID: ${item.conversationId}`);
    }

    // Get categories - this IS async
    if (item.categories && typeof item.categories.getAsync === 'function') {
      item.categories.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const categories = result.value && result.value.length > 0 ? result.value.join(', ') : 'None';
          properties.push(`Categories: ${categories}`);
          updateResultArea(properties);
        } else {
          properties.push(`Categories: Unable to load`);
          updateResultArea(properties);
        }
      });
    } else {
      properties.push(`Categories: Not available`);
      updateResultArea(properties);
    }

    // Get body preview - this IS async
    if (item.body && typeof item.body.getAsync === 'function') {
      item.body.getAsync(Office.CoercionType.Text, { asyncContext: properties }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const bodyPreview = result.value.substring(0, 200) + '...';
          result.asyncContext.push(`Body Preview: ${bodyPreview}`);
          updateResultArea(result.asyncContext);
        } else {
          result.asyncContext.push(`Body Preview: Unable to load`);
          updateResultArea(result.asyncContext);
        }
      });
    } else {
      properties.push(`Body Preview: Not available`);
      updateResultArea(properties);
    }

    // Add attachment info
    if (item.attachments && item.attachments.length > 0) {
      const attachmentInfo = item.attachments.map(att => 
        `${att.name} (${att.attachmentType})`
      ).join(', ');
      properties.push(`Attachments: ${attachmentInfo}`);
    } else {
      properties.push(`Attachments: None`);
    }

  } catch (error) {
    console.error("Error getting item properties:", error);
    updateResultArea([`Error: ${error.message}`]);
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