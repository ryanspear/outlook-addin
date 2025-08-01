/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * Enhanced Outlook Add-in for Loan Application Email Parsing
 */

/* global document, Office */

// Import polyfills at the top
import "core-js/stable";
import "regenerator-runtime/runtime";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("get-properties-button").onclick = getItemProperties;
    document.getElementById("parse-email-button").onclick = parseEmailContent;
    
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
 * Parse email content for loan applicant, property, and company details
 */
function parseEmailContent() {
  showParsingStatus("loading", "Analyzing email content...");
  
  try {
    const item = Office.context.mailbox.item;
    
    if (!item.body || typeof item.body.getAsync !== 'function') {
      showParsingStatus("error", "Unable to access email body");
      return;
    }

    // Get email body content
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const emailContent = result.value;
        
        // Also get subject for additional context
        const subject = item.subject || "";
        const fullContent = subject + "\n\n" + emailContent;
        
        // Extract information
        const extractedData = extractLoanInformation(fullContent);
        
        // Display results
        displayExtractedInformation(extractedData);
        showParsingStatus("success", "Email analysis complete!");
        
      } else {
        showParsingStatus("error", "Failed to retrieve email content");
        console.error("Error getting email body:", result.error);
      }
    });
    
  } catch (error) {
    console.error("Error parsing email content:", error);
    showParsingStatus("error", `Error: ${error.message}`);
  }
}

/**
 * Extract loan-related information from email content
 * @param {string} content - Email content to analyze
 * @returns {Object} Extracted information object
 */
function extractLoanInformation(content) {
  const extractedData = {
    applicant: {},
    property: {},
    company: {}
  };

  // Convert to lowercase for easier matching (but keep original for display)
  const lowerContent = content.toLowerCase();
  
  // Extract Applicant Information
  extractedData.applicant = extractApplicantDetails(content, lowerContent);
  
  // Extract Property Information
  extractedData.property = extractPropertyDetails(content, lowerContent);
  
  // Extract Company Information
  extractedData.company = extractCompanyDetails(content, lowerContent);
  
  return extractedData;
}

/**
 * Extract loan applicant details
 * @param {string} content - Original content
 * @param {string} lowerContent - Lowercase content for matching
 * @returns {Object} Applicant information
 */
function extractApplicantDetails(content, lowerContent) {
  const applicant = {};
  
  // Email addresses - comprehensive regex
  const emailRegex = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
  const emails = content.match(emailRegex);
  if (emails && emails.length > 0) {
    applicant.emails = [...new Set(emails)]; // Remove duplicates
  }
  
  // Phone numbers - UK and international formats
  const phoneRegex = /(?:\+44\s?|0)(?:\d{2,4}\s?\d{3,4}\s?\d{3,4}|\d{3}\s?\d{3}\s?\d{4}|\d{4}\s?\d{6})/g;
  const phones = content.match(phoneRegex);
  if (phones && phones.length > 0) {
    applicant.phones = [...new Set(phones)];
  }
  
  // Names - look for common patterns
  const namePatterns = [
    /(?:applicant[:\s]+|client[:\s]+|borrower[:\s]+|mr\.?\s+|mrs\.?\s+|ms\.?\s+|miss\s+)([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)/gi,
    /(?:name[:\s]+)([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)/gi,
    /(?:dear\s+)([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)/gi
  ];
  
  const names = [];
  namePatterns.forEach(pattern => {
    const matches = content.match(pattern);
    if (matches) {
      matches.forEach(match => {
        const name = match.replace(/^(?:applicant[:\s]+|client[:\s]+|borrower[:\s]+|mr\.?\s+|mrs\.?\s+|ms\.?\s+|miss\s+|name[:\s]+|dear\s+)/gi, '').trim();
        if (name.length > 2 && name.split(' ').length >= 2) {
          names.push(name);
        }
      });
    }
  });
  
  if (names.length > 0) {
    applicant.names = [...new Set(names)];
  }
  
  // Income information
  const incomeRegex = /(?:income|salary|earnings?)[:\s]*£?([0-9,]+(?:\.[0-9]{2})?)/gi;
  const incomeMatches = content.match(incomeRegex);
  if (incomeMatches) {
    applicant.income = incomeMatches;
  }
  
  // Employment information
  const employmentRegex = /(?:employer|employed by|works? (?:at|for))[:\s]*([A-Za-z0-9\s&.,'-]+?)(?=\n|\.|,|$)/gi;
  const employmentMatches = content.match(employmentRegex);
  if (employmentMatches) {
    applicant.employment = employmentMatches.map(match => 
      match.replace(/^(?:employer|employed by|works? (?:at|for))[:\s]*/gi, '').trim()
    );
  }
  
  return applicant;
}

/**
 * Extract property details
 * @param {string} content - Original content
 * @param {string} lowerContent - Lowercase content for matching
 * @returns {Object} Property information
 */
function extractPropertyDetails(content, lowerContent) {
  const property = {};
  
  // UK Postcodes
  const postcodeRegex = /\b[A-Z]{1,2}[0-9][A-Z0-9]?\s?[0-9][A-Z]{2}\b/g;
  const postcodes = content.match(postcodeRegex);
  if (postcodes && postcodes.length > 0) {
    property.postcodes = [...new Set(postcodes)];
  }
  
  // Property addresses - look for common address patterns
  const addressPatterns = [
    /(?:property|address|located at)[:\s]*([0-9]+[A-Za-z]?\s+[A-Za-z\s,'.-]+(?:[A-Z]{1,2}[0-9][A-Z0-9]?\s?[0-9][A-Z]{2})?)/gi,
    /([0-9]+[A-Za-z]?\s+[A-Za-z\s,'.-]+[A-Z]{1,2}[0-9][A-Z0-9]?\s?[0-9][A-Z]{2})/g
  ];
  
  const addresses = [];
  addressPatterns.forEach(pattern => {
    const matches = content.match(pattern);
    if (matches) {
      matches.forEach(match => {
        const address = match.replace(/^(?:property|address|located at)[:\s]*/gi, '').trim();
        if (address.length > 10) {
          addresses.push(address);
        }
      });
    }
  });
  
  if (addresses.length > 0) {
    property.addresses = [...new Set(addresses)];
  }
  
  // Property value/price
  const valueRegex = /(?:value|price|purchase price|valuation)[:\s]*£([0-9,]+(?:\,[0-9]{3})*(?:\.[0-9]{2})?)/gi;
  const valueMatches = content.match(valueRegex);
  if (valueMatches) {
    property.values = valueMatches;
  }
  
  // Property type
  const typeRegex = /(?:property type|type of property)[:\s]*([A-Za-z\s-]+?)(?=\n|\.|,|$)/gi;
  const typeMatches = content.match(typeRegex);
  if (typeMatches) {
    property.types = typeMatches.map(match => 
      match.replace(/^(?:property type|type of property)[:\s]*/gi, '').trim()
    );
  }
  
  // Look for common property types in context
  const propertyTypes = ['flat', 'apartment', 'house', 'bungalow', 'cottage', 'mansion', 'maisonette', 'studio'];
  const foundTypes = [];
  propertyTypes.forEach(type => {
    if (lowerContent.includes(type)) {
      foundTypes.push(type);
    }
  });
  
  if (foundTypes.length > 0) {
    property.detectedTypes = foundTypes;
  }
  
  return property;
}

/**
 * Extract company details
 * @param {string} content - Original content
 * @param {string} lowerContent - Lowercase content for matching
 * @returns {Object} Company information
 */
function extractCompanyDetails(content, lowerContent) {
  const company = {};
  
  // Company registration numbers
  const companyNumberRegex = /(?:company number|registration number|reg\.?\s*no\.?|companies house)[:\s]*([0-9]{8}|[A-Z]{2}[0-9]{6})/gi;
  const companyNumbers = content.match(companyNumberRegex);
  if (companyNumbers && companyNumbers.length > 0) {
    company.registrationNumbers = companyNumbers.map(match => 
      match.replace(/^(?:company number|registration number|reg\.?\s*no\.?|companies house)[:\s]*/gi, '').trim()
    );
  }
  
  // Company names - look for Ltd, Limited, PLC, etc.
  const companyNameRegex = /([A-Za-z0-9\s&'.-]+?\s+(?:Ltd|Limited|PLC|plc|LLP|Partnership|Company|Corp|Corporation|Inc|Incorporated)\.?)/g;
  const companyNames = content.match(companyNameRegex);
  if (companyNames && companyNames.length > 0) {
    company.names = [...new Set(companyNames.map(name => name.trim()))];
  }
  
  // Trading as / DBA
  const tradingAsRegex = /(?:trading as|t\/a|dba)[:\s]*([A-Za-z0-9\s&'.-]+?)(?=\n|\.|,|$)/gi;
  const tradingAsMatches = content.match(tradingAsRegex);
  if (tradingAsMatches) {
    company.tradingAs = tradingAsMatches.map(match => 
      match.replace(/^(?:trading as|t\/a|dba)[:\s]*/gi, '').trim()
    );
  }
  
  // VAT numbers
  const vatRegex = /(?:vat number|vat reg)[:\s]*(?:gb\s?)?([0-9]{9})/gi;
  const vatNumbers = content.match(vatRegex);
  if (vatNumbers) {
    company.vatNumbers = vatNumbers.map(match => 
      match.replace(/^(?:vat number|vat reg)[:\s]*(?:gb\s?)?/gi, '').trim()
    );
  }
  
  return company;
}

/**
 * Display extracted information in the UI
 * @param {Object} data - Extracted data object
 */
function displayExtractedInformation(data) {
  const extractedInfoDiv = document.getElementById("extracted-info");
  extractedInfoDiv.style.display = "block";
  
  // Display applicant information
  displayApplicantInfo(data.applicant);
  
  // Display property information
  displayPropertyInfo(data.property);
  
  // Display company information
  displayCompanyInfo(data.company);
}

/**
 * Display applicant information
 * @param {Object} applicant - Applicant data
 */
function displayApplicantInfo(applicant) {
  const applicantDiv = document.getElementById("applicant-info");
  let html = "";
  
  if (applicant.names && applicant.names.length > 0) {
    html += `<div class="info-item"><strong>Names:</strong> ${applicant.names.join(', ')}</div>`;
  }
  
  if (applicant.emails && applicant.emails.length > 0) {
    html += `<div class="info-item"><strong>Email Addresses:</strong> ${applicant.emails.join(', ')}</div>`;
  }
  
  if (applicant.phones && applicant.phones.length > 0) {
    html += `<div class="info-item"><strong>Phone Numbers:</strong> ${applicant.phones.join(', ')}</div>`;
  }
  
  if (applicant.income && applicant.income.length > 0) {
    html += `<div class="info-item"><strong>Income Information:</strong> ${applicant.income.join(', ')}</div>`;
  }
  
  if (applicant.employment && applicant.employment.length > 0) {
    html += `<div class="info-item"><strong>Employment:</strong> ${applicant.employment.join(', ')}</div>`;
  }
  
  if (html === "") {
    html = '<div class="info-item not-found">No applicant information found</div>';
  }
  
  applicantDiv.innerHTML = html;
}

/**
 * Display property information
 * @param {Object} property - Property data
 */
function displayPropertyInfo(property) {
  const propertyDiv = document.getElementById("property-info");
  let html = "";
  
  if (property.addresses && property.addresses.length > 0) {
    html += `<div class="info-item"><strong>Addresses:</strong> ${property.addresses.join('<br>')}</div>`;
  }
  
  if (property.postcodes && property.postcodes.length > 0) {
    html += `<div class="info-item"><strong>Postcodes:</strong> ${property.postcodes.join(', ')}</div>`;
  }
  
  if (property.values && property.values.length > 0) {
    html += `<div class="info-item"><strong>Property Values:</strong> ${property.values.join(', ')}</div>`;
  }
  
  if (property.types && property.types.length > 0) {
    html += `<div class="info-item"><strong>Property Types:</strong> ${property.types.join(', ')}</div>`;
  }
  
  if (property.detectedTypes && property.detectedTypes.length > 0) {
    html += `<div class="info-item"><strong>Detected Types:</strong> ${property.detectedTypes.join(', ')}</div>`;
  }
  
  if (html === "") {
    html = '<div class="info-item not-found">No property information found</div>';
  }
  
  propertyDiv.innerHTML = html;
}

/**
 * Display company information
 * @param {Object} company - Company data
 */
function displayCompanyInfo(company) {
  const companyDiv = document.getElementById("company-info");
  let html = "";
  
  if (company.names && company.names.length > 0) {
    html += `<div class="info-item"><strong>Company Names:</strong> ${company.names.join(', ')}</div>`;
  }
  
  if (company.registrationNumbers && company.registrationNumbers.length > 0) {
    company.registrationNumbers.forEach(regNumber => {
      html += `<div class="info-item"><strong>Registration Number:</strong> ${regNumber}`;
      html += `<br><a href="https://find-and-update.company-information.service.gov.uk/company/${regNumber}" target="_blank" class="companies-house-link">🔍 Search Companies House</a></div>`;
    });
  }
  
  if (company.tradingAs && company.tradingAs.length > 0) {
    html += `<div class="info-item"><strong>Trading As:</strong> ${company.tradingAs.join(', ')}</div>`;
  }
  
  if (company.vatNumbers && company.vatNumbers.length > 0) {
    html += `<div class="info-item"><strong>VAT Numbers:</strong> ${company.vatNumbers.join(', ')}</div>`;
  }
  
  if (html === "") {
    html = '<div class="info-item not-found">No company information found</div>';
  }
  
  companyDiv.innerHTML = html;
}

/**
 * Show parsing status to user
 * @param {string} type - Status type: loading, success, error
 * @param {string} message - Status message
 */
function showParsingStatus(type, message) {
  const statusDiv = document.getElementById("parsing-status");
  statusDiv.className = `parsing-status ${type}`;
  statusDiv.textContent = message;
  statusDiv.style.display = "block";
  
  if (type === "success" || type === "error") {
    setTimeout(() => {
      statusDiv.style.display = "none";
    }, 3000);
  }
}

/**
 * Get detailed email properties (original functionality)
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

    // Show results area
    const resultArea = document.getElementById("result-area");
    resultArea.style.display = "block";

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