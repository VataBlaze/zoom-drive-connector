/**
 * Zoom to Google Drive Connector Template
 * 
 * This Google Apps Script transfers Zoom cloud recordings and AI meeting summaries 
 * to Google Drive and tracks transfers in a Google Sheet.
 * 
 * SETUP INSTRUCTIONS:
 * 1. Replace all placeholder values in the CONFIG object below with your actual credentials
 * 2. Set up a Zoom Server-to-Server OAuth app in the Zoom Marketplace
 * 3. Create folders in Google Drive and get their folder IDs
 * 4. Run the setup() function once to initialize the tracking sheet and daily trigger
 * 5. Test by running runZoomToDriveConnector() manually
 */

const CONFIG = {
  // Zoom API credentials - Get these from your Zoom Server-to-Server OAuth app
  ZOOM_ACCOUNT_ID: "YOUR_ZOOM_ACCOUNT_ID",
  ZOOM_CLIENT_ID: "YOUR_ZOOM_CLIENT_ID", 
  ZOOM_CLIENT_SECRET: "YOUR_ZOOM_CLIENT_SECRET",
  
  // Google Drive destination folder ID - Right-click folder in Drive > Get link > Extract ID from URL
  GOOGLE_DRIVE_FOLDER_ID: "YOUR_MAIN_DRIVE_FOLDER_ID",
  
  // Date settings - How many days back to fetch recordings
  DAYS_TO_FETCH: 2,
  
  // Processing settings - Set to false if you want to keep recordings in Zoom
  DELETE_AFTER_TRANSFER: true,
  
  // Tracking sheet - Leave empty to auto-create, or provide existing sheet ID
  TRACKING_SHEET_ID: "", // Will be auto-populated if left empty
  TRACKING_SHEET_NAME: "Transfer Tracker",

  // Folder mapping - Maps meeting topic keywords to specific Drive folder IDs
  // Add your own keywords and corresponding folder IDs
  FOLDER_MAPPING: {
    "operations": "YOUR_OPERATIONS_FOLDER_ID",
    "council": "YOUR_COUNCIL_FOLDER_ID", 
    "marketing": "YOUR_MARKETING_FOLDER_ID",
    "programming": "YOUR_PROGRAMMING_FOLDER_ID"
    // Add more keyword mappings as needed
  },
  
  // Default folder ID for recordings that don't match any keywords
  DEFAULT_FOLDER_ID: "YOUR_DEFAULT_FOLDER_ID",
  
  // Meeting topic prefix to remove (optional)
  TOPIC_PREFIX_TO_REMOVE: "Your Organization Name" // Set to "" to disable
};

function resetZoomCredentials() {
  // Clear any cached tokens and force new authentication
  PropertiesService.getScriptProperties().deleteProperty('zoom_token');
  Logger.log("Zoom credentials have been reset. Run the connector again to generate a new token.");
}

/**
 * Main function to run the connector
 */
function runZoomToDriveConnector() {
  Logger.log("Starting Zoom to Google Drive connector");
  
  try {
    // Check/create tracking sheet
    const trackingSheet = initTrackingSheet();
    
    // Get Zoom access token
    const zoomToken = getZoomAccessToken();
    if (!zoomToken) {
      throw new Error("Failed to get Zoom access token");
    }
    
    // Get all Zoom users in the account 
    const users = getAllZoomUsers(zoomToken);
    Logger.log(`Found ${users.length} users in Zoom account`);
    
    // Process recordings for each user
    for (const user of users) {
      Logger.log(`Processing recordings for user: ${user.email}`);
      const recordings = getZoomRecordingsForUser(zoomToken, user.id);
      Logger.log(`Found ${recordings.length} recordings for user ${user.email}`);
      
      // Process each recording
      for (const recording of recordings) {
        processRecording(recording, zoomToken, trackingSheet);
      }
    }

    // Process summaries for each user
    for (const user of users) {
      Logger.log(`Processing summaries for user: ${user.email}`);
      const summaries = getZoomSummariesForUser(zoomToken, user.id);
      Logger.log(`Found ${summaries.length} summaries for user ${user.email}`);
    
      // Process each meeting summary
      for (const meeting of summaries) {
        processMeetingSummary(meeting, zoomToken, trackingSheet);
      }
    }

    Logger.log("Finished processing all recordings and transcripts");
  } catch (error) {
    Logger.log(`Error running connector: ${error.message}`);
    trackError(error);
  }
}

/**
 * Get all users in the Zoom account
 */
function getAllZoomUsers(accessToken) {
  try {
    // Use the list users endpoint
    const usersUrl = "https://api.zoom.us/v2/users";
    const options = {
      method: "get",
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      },
      muteHttpExceptions: true
    };
    
    let allUsers = [];
    let hasMorePages = true;
    let pageNumber = 1;
    
    while (hasMorePages) {
      const url = `${usersUrl}?page_size=300&page_number=${pageNumber}&status=active`;
      Logger.log(`Fetching users page ${pageNumber}`);
      
      const response = UrlFetchApp.fetch(url, options);
      
      if (response.getResponseCode() === 200) {
        const data = JSON.parse(response.getContentText());
        
        if (data.users && data.users.length > 0) {
          allUsers = allUsers.concat(data.users);
          
          // Log users found
          data.users.forEach(user => {
            Logger.log(`Found user: ${user.email} (${user.id})`);
          });
          
          // Check if there are more pages
          if (data.page_count > pageNumber) {
            pageNumber++;
          } else {
            hasMorePages = false;
          }
        } else {
          hasMorePages = false;
        }
      } else {
        Logger.log(`Error fetching users: ${response.getContentText()}`);
        hasMorePages = false;
      }
      
      // Add a small delay to avoid rate limiting
      if (hasMorePages) {
        Utilities.sleep(500);
      }
    }
    
    return allUsers;
  } catch (error) {
    Logger.log(`Error getting Zoom users: ${error.message}`);
    return [];
  }
}

/**
 * Initialize or get the tracking spreadsheet
 */
function initTrackingSheet() {
  let spreadsheet;
  let sheet;
  
  // If no tracking sheet ID is specified, create a new one
  if (!CONFIG.TRACKING_SHEET_ID) {
    Logger.log("Creating new tracking spreadsheet");
    spreadsheet = SpreadsheetApp.create(CONFIG.TRACKING_SHEET_NAME);
    CONFIG.TRACKING_SHEET_ID = spreadsheet.getId();
    
    // Log the new spreadsheet ID so it can be saved in the config
    Logger.log(`Created new tracking spreadsheet with ID: ${CONFIG.TRACKING_SHEET_ID}`);
    Logger.log(`IMPORTANT: Update your CONFIG.TRACKING_SHEET_ID with this value: ${CONFIG.TRACKING_SHEET_ID}`);
    
    // Get the first sheet
    sheet = spreadsheet.getSheets()[0];
    
    // Set up the header row
    sheet.getRange("A1:F1").setValues([[
      "Meeting Topic", 
      "Meeting Date", 
      "Files Transferred",
      "Location in Drive", 
      "Transfer Date",
      "Host Email"
    ]]);
    
    // Format header row
    sheet.getRange("A1:F1").setFontWeight("bold");
    sheet.setFrozenRows(1);
  } else {
    // Open existing spreadsheet
    spreadsheet = SpreadsheetApp.openById(CONFIG.TRACKING_SHEET_ID);
    sheet = spreadsheet.getSheetByName(CONFIG.TRACKING_SHEET_NAME);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = spreadsheet.insertSheet(CONFIG.TRACKING_SHEET_NAME);
      // Set up the header row
      sheet.getRange("A1:F1").setValues([[
        "Meeting Topic", 
        "Meeting Date", 
        "Files Transferred",
        "Location in Drive", 
        "Transfer Date",
        "Host Email"
      ]]);
      sheet.getRange("A1:F1").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
    
    // Add Host Email column if it doesn't exist
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.includes("Host Email")) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue("Host Email");
    }
    
    // Set date format for Meeting Date column
    sheet.getRange("B:B").setNumberFormat("MM-dd-yyyy");
    
    // Set date format for Transfer Date column
    sheet.getRange("E:E").setNumberFormat("MM-dd-yyyy");
  }
  
  return sheet;
}

/**
 * Get Zoom OAuth token using client credentials grant
 */
function getZoomAccessToken() {
  const tokenUrl = "https://zoom.us/oauth/token";
  const authHeader = Utilities.base64Encode(`${CONFIG.ZOOM_CLIENT_ID}:${CONFIG.ZOOM_CLIENT_SECRET}`);
  
  const options = {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    headers: {
      "Authorization": `Basic ${authHeader}`
    },
    payload: {
      "grant_type": "account_credentials",
      "account_id": CONFIG.ZOOM_ACCOUNT_ID
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(tokenUrl, options);
    const responseData = JSON.parse(response.getContentText());
    
    if (response.getResponseCode() === 200 && responseData.access_token) {
      return responseData.access_token;
    } else {
      Logger.log(`Failed to get Zoom token: ${response.getContentText()}`);
      return null;
    }
  } catch (error) {
    Logger.log(`Error getting Zoom access token: ${error.message}`);
    return null;
  }
}

/**
 * Get cloud recordings for a specific user from Zoom API
 */
function getZoomRecordingsForUser(accessToken, userId) {
  try {
    // Calculate date range
    const today = new Date();
    const fromDate = new Date(today);
    fromDate.setDate(today.getDate() - CONFIG.DAYS_TO_FETCH);
    
    const fromDateStr = Utilities.formatDate(fromDate, "GMT", "yyyy-MM-dd");
    const toDateStr = Utilities.formatDate(today, "GMT", "yyyy-MM-dd");
    
    Logger.log(`Fetching recordings for user ${userId} from ${fromDateStr} to ${toDateStr} (${CONFIG.DAYS_TO_FETCH} days)`);
    
    // The API only allows 30 days of data per request, so we need to make multiple requests
    const chunkSize = 30; // days per request
    let allRecordings = [];
    
    // Calculate how many chunks we need
    const numChunks = Math.ceil(CONFIG.DAYS_TO_FETCH / chunkSize);
    
    for (let i = 0; i < numChunks; i++) {
      // Calculate start and end date for this chunk
      const chunkEndDate = new Date(today);
      chunkEndDate.setDate(today.getDate() - (i * chunkSize));
      
      const chunkStartDate = new Date(chunkEndDate);
      chunkStartDate.setDate(chunkEndDate.getDate() - chunkSize);
      
      // Make sure we don't go beyond the requested fromDate
      if (chunkStartDate < fromDate) {
        chunkStartDate.setTime(fromDate.getTime());
      }
      
      // Format dates for API
      const chunkStartStr = Utilities.formatDate(chunkStartDate, "GMT", "yyyy-MM-dd");
      const chunkEndStr = Utilities.formatDate(chunkEndDate, "GMT", "yyyy-MM-dd");
      
      Logger.log(`Fetching chunk ${i+1}/${numChunks} for user ${userId}: ${chunkStartStr} to ${chunkEndStr}`);
      
      // Get recordings for this date range
      const chunkRecordings = fetchRecordingsForDateRange(accessToken, userId, chunkStartStr, chunkEndStr);
      allRecordings = allRecordings.concat(chunkRecordings);
      
      // Add a small delay to avoid rate limiting
      if (i < numChunks - 1) {
        Utilities.sleep(500);
      }
    }
    
    Logger.log(`Total recordings found for user ${userId}: ${allRecordings.length}`);
    return allRecordings;
    
  } catch (error) {
    Logger.log(`Error in getZoomRecordingsForUser: ${error.message}`);
    return [];
  }
}

/**
 * Get summaries for a specific user from Zoom API
 */
function getZoomSummariesForUser(accessToken, userId) {
  try {
    // Calculate date range
    const today = new Date();
    const fromDate = new Date(today);
    fromDate.setDate(today.getDate() - CONFIG.DAYS_TO_FETCH);
    
    const fromDateStr = Utilities.formatDate(fromDate, "GMT", "yyyy-MM-dd");
    const toDateStr = Utilities.formatDate(today, "GMT", "yyyy-MM-dd");
    
    Logger.log(`Fetching summaries for user ${userId} from ${fromDateStr} to ${toDateStr} (${CONFIG.DAYS_TO_FETCH} days)`);
    
    // The API only allows 30 days of data per request, so we need to make multiple requests
    const chunkSize = 30; // days per request
    let allSummaries = [];
    
    // Calculate how many chunks we need
    const numChunks = Math.ceil(CONFIG.DAYS_TO_FETCH / chunkSize);
    
    for (let i = 0; i < numChunks; i++) {
      // Calculate start and end date for this chunk
      const chunkEndDate = new Date(today);
      chunkEndDate.setDate(today.getDate() - (i * chunkSize));
      
      const chunkStartDate = new Date(chunkEndDate);
      chunkStartDate.setDate(chunkEndDate.getDate() - chunkSize);
      
      // Make sure we don't go beyond the requested fromDate
      if (chunkStartDate < fromDate) {
        chunkStartDate.setTime(fromDate.getTime());
      }
      
      // Format dates for API
      const chunkStartStr = Utilities.formatDate(chunkStartDate, "GMT", "yyyy-MM-dd'T'00:00:00'Z'");
      const chunkEndStr   = Utilities.formatDate(chunkEndDate,   "GMT", "yyyy-MM-dd'T'23:59:59'Z'");
      
      Logger.log(
        `Fetching chunk ${i+1}/${numChunks} for user ${userId}: ` +
        `${chunkStartStr} to ${chunkEndStr}`
      );
      
      // Get recordings for this date range
      const chunkRecordings = fetchSummariesForDateRange(accessToken, chunkStartStr, chunkEndStr);
      allSummaries = allSummaries.concat(chunkRecordings);
      
      // Add a small delay to avoid rate limiting
      if (i < numChunks - 1) {
        Utilities.sleep(500);
      }
    }
    
    Logger.log(`Total recordings found for user ${userId}: ${allSummaries.length}`);
    return allSummaries;
    
  } catch (error) {
    Logger.log(`Error in getZoomSummariesForUser: ${error.message}`);
    return [];
  }
}

function getMeetingSummaryDetail(meetingId, accessToken) {
  Logger.log(`Starting getMeetingSummaryDetail for meetingId: ${meetingId}`);
  
  const url = `https://api.zoom.us/v2/meetings/${meetingId}/meeting_summary`;
  
  const options = {
    method: "get",
    headers: {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    muteHttpExceptions: true
  };
  
  try {
    const resp = UrlFetchApp.fetch(url, options);
    const statusCode = resp.getResponseCode();
    
    const responseText = resp.getContentText();
    
    if (statusCode === 200) {
      const data = JSON.parse(responseText);
      
      // Extract relevant fields from the schema
      const result = {
        meetingId: data.meeting_id,
        meetingTopic: data.meeting_topic,
        overviewSummary: data.summary_overview,
        detailedSummary: data.summary_details,
        meetingDate: data.summary_created_time
      };
      return result;
    } else {
      Logger.log(`Failed to fetch detail for meeting ${meetingId}. Response: ${responseText}`);
      return null;
    }
  } catch (error) {
    Logger.log(`Error fetching meeting summary for ${meetingId}: ${error.message}`);
    return null;
  }
}

/**
 * Fetch recordings for a specific date range
 */
function fetchRecordingsForDateRange(accessToken, userId, fromDate, toDate) {
  const apiUrl = `https://api.zoom.us/v2/users/${userId}/recordings`;
  const options = {
    method: "get",
    headers: {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    muteHttpExceptions: true
  };
  
  let recordings = [];
  let page = 1;
  let hasMorePages = true;
  
  while (hasMorePages) {
    const url = `${apiUrl}?from=${fromDate}&to=${toDate}&page_size=100&page_number=${page}`;
    Logger.log(`Fetching page ${page} for user ${userId}, date range ${fromDate} to ${toDate}...`);
    
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() === 200) {
      const responseData = JSON.parse(response.getContentText());
      
      // Debug output
      Logger.log(`API response for ${userId}, ${fromDate} to ${toDate}: page ${page}, found ${responseData.meetings?.length || 0} recordings`);
      
      if (responseData.meetings && responseData.meetings.length > 0) {
        // Add host_email to each recording for tracking
        responseData.meetings.forEach(meeting => {
          meeting.host_email = meeting.host_email || getUserEmailById(userId, accessToken);
        });
        
        recordings = recordings.concat(responseData.meetings);
        
        // Log each recording found
        responseData.meetings.forEach(meeting => {
          Logger.log(`Found recording: ${meeting.topic} (${meeting.start_time}) - Host: ${meeting.host_email}`);
        });
      }
      
      // Check if there are more pages
      if (responseData.page_count > page) {
        page++;
      } else {
        hasMorePages = false;
      }
    } else {
      Logger.log(`API error for dates ${fromDate} to ${toDate}: ${response.getContentText()}`);
      hasMorePages = false;
    }
  }
  
  return recordings;
}

/**
 * Fetch summaries for a specific date range
 */
function fetchSummariesForDateRange(accessToken, fromDate, toDate) {
  const apiUrl = 'https://api.zoom.us/v2/meetings/meeting_summaries';
  const options = {
    method: "get",
    headers: {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    muteHttpExceptions: true
  };

  let summaries = [];
  let hasMorePages = true;
  let nextPageToken = '';

  while (hasMorePages) {
    let url = `${apiUrl}?from=${encodeURIComponent(fromDate)}&to=${encodeURIComponent(toDate)}&page_size=100`;
    if (nextPageToken) {
      url += `&next_page_token=${nextPageToken}`;
    }

    Logger.log(`Fetching meeting summaries for date range ${fromDate} to ${toDate}...`);

    try {
      const response = UrlFetchApp.fetch(url, options);

      if (response.getResponseCode() === 200) {
        const responseData = JSON.parse(response.getContentText());

        Logger.log(`API response: Found ${responseData.summaries?.length || 0} summaries`);
        if (responseData.summaries && responseData.summaries.length > 0) {
          summaries = summaries.concat(responseData.summaries);
        }

        // Check if there are more pages
        if (responseData.next_page_token) {
          nextPageToken = responseData.next_page_token;
        } else {
          hasMorePages = false;
        }
      } else {
        Logger.log(`API error: ${response.getContentText()}`);
        hasMorePages = false;
      }
    } catch (error) {
      Logger.log(`Error fetching summaries: ${error.message}`);
      hasMorePages = false;
    }
  }

  return summaries;
}

/**
 * Get user email by user ID (for when the API response doesn't include it)
 */
function getUserEmailById(userId, accessToken) {
  const apiUrl = `https://api.zoom.us/v2/users/${userId}`;
  const options = {
    method: "get",
    headers: {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    
    if (response.getResponseCode() === 200) {
      const userData = JSON.parse(response.getContentText());
      return userData.email || "Unknown Email";
    } else {
      return "Unknown Email";
    }
  } catch (error) {
    Logger.log(`Error getting user email for ID ${userId}: ${error.message}`);
    return "Unknown Email";
  }
}

/**
 * Convert VTT transcript to plain text and save as TXT file
 */
function convertTranscriptToText(vttContent, fileName, folderId) {
  // Split VTT content into lines
  let lines = vttContent.split("\n");
  let plainText = [];
  let currentSpeaker = null;
  let currentStatement = [];
  
  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trim();
    
    // Skip WebVTT header, timestamps, sequence numbers, and empty lines
    if (line.match(/^WEBVTT/) || 
        line.match(/^\d{2}:\d{2}:\d{2}/) || 
        line.match(/^\d+$/) || 
        line === "") {
      continue;
    }
    
    // Check if line starts with a speaker (format: "Speaker: Text")
    if (line.includes(":")) {
      let [speaker, ...text] = line.split(":");
      speaker = speaker.trim();
      let statement = text.join(":").trim();
      
      if (speaker && statement) {
        if (speaker === currentSpeaker) {
          // Same speaker, append to current statement
          currentStatement.push(statement);
        } else {
          // New speaker, save previous statement if exists
          if (currentSpeaker && currentStatement.length > 0) {
            plainText.push(`${currentSpeaker}: ${currentStatement.join(" ")}`);
          }
          // Update to new speaker
          currentSpeaker = speaker;
          currentStatement = [statement];
        }
      }
    }
  }
  
  // Add the last speaker's statement
  if (currentSpeaker && currentStatement.length > 0) {
    plainText.push(`${currentSpeaker}: ${currentStatement.join(" ")}`);
  }
  
  // Join statements with a single empty line
  let finalText = plainText.join("\n\n");
  
  // Create a plain text file with the transcript content
  const textFileName = fileName.replace(".vtt", ".txt").replace(".VTT", ".txt");
  const folder = DriveApp.getFolderById(folderId);
  
  // Create the text file in Drive
  const textFile = folder.createFile(textFileName, finalText, MimeType.PLAIN_TEXT);
  
  return textFile;
}

/**
 * Process a single recording: download, upload to Drive, delete from Zoom
 */
function processRecording(recording, accessToken, trackingSheet) {
  try {
    const meetingId = recording.uuid;
    let meetingTopic = recording.topic || "Unnamed Meeting";
    
    // Remove topic prefix if configured
    if (CONFIG.TOPIC_PREFIX_TO_REMOVE && meetingTopic.startsWith(CONFIG.TOPIC_PREFIX_TO_REMOVE)) {
      const prefixPattern = new RegExp(`^${CONFIG.TOPIC_PREFIX_TO_REMOVE}(\\s\\|\\s|\\s)`);
      meetingTopic = meetingTopic.replace(prefixPattern, "");
    }
    
    const meetingDate = recording.start_time ? recording.start_time.split("T")[0] : "Unknown Date";
    const hostEmail = recording.host_email || "Unknown Host";
    
    // Check if recording has already been processed
    if (isRecordingProcessed(meetingTopic, meetingDate, trackingSheet)) {
      Logger.log(`Recording ${meetingTopic} (${meetingId}) - Host: ${hostEmail} already processed. Skipping.`);
      return;
    }
    
    Logger.log(`Processing recording: ${meetingTopic} (${meetingId}) - Host: ${hostEmail}`);
    
    // Create a folder in Google Drive for this meeting
    const folderName = `${meetingDate}_${meetingTopic}`;
    const destinationFolderId = getDestinationFolderId(meetingTopic);
    const folder = createDriveFolder(folderName, destinationFolderId);
    
    let filesTransferred = [];
    
    // Process each recording file
    if (recording.recording_files && recording.recording_files.length > 0) {
      for (const file of recording.recording_files) {
        const fileType = file.file_type?.toLowerCase();
        const fileExtension = file.file_extension?.toLowerCase();

        // Skip JSON timeline files
        if (fileExtension === 'json') {
          continue;
        }

        const fileName = fileExtension === 'mp4'
          ? `${meetingDate}_${meetingTopic}_video.${fileExtension}`
          : fileExtension === 'm4a'
          ? `${meetingDate}_${meetingTopic}_audio.${fileExtension}`
          : fileType === 'transcript' 
          ? `${meetingDate}_${meetingTopic}_transcript.${fileExtension}`
          : `${meetingDate}_${meetingTopic}_${fileType}.${fileExtension}`;
        const downloadUrl = file.download_url;
        
        if (downloadUrl) {
          // Add access token to URL
          const authDownloadUrl = downloadUrl + (downloadUrl.includes('?') ? '&' : '?') + 
                                `access_token=${accessToken}`;
          
          try {
            // Special handling for transcript files
            if (fileExtension === 'vtt' || fileExtension === 'VTT') {
              // Download the VTT content
              const response = UrlFetchApp.fetch(authDownloadUrl, {
                method: "get",
                muteHttpExceptions: true,
                followRedirects: true
              });
              
              if (response.getResponseCode() === 200) {
                const vttContent = response.getContentText();
                const textFile = convertTranscriptToText(vttContent, fileName, folder.getId());
                const textFileName = fileName.replace(".vtt", ".txt").replace(".VTT", ".txt");
                filesTransferred.push(`${textFileName} (${textFile.getId()})`);
              }
            } else {
              // Regular file download and upload
              const driveFile = downloadAndUploadFile(authDownloadUrl, fileName, folder.getId(), getMimeType(fileExtension));
              
              if (driveFile) {
                filesTransferred.push(`${fileName} (${driveFile.getId()})`);
              }
            }
          } catch (error) {
            Logger.log(`Error processing file ${fileName}: ${error.message}`);
          }
        }
      }
    }
    
    // Delete from Zoom if configured and successful transfer
    if (CONFIG.DELETE_AFTER_TRANSFER && filesTransferred.length > 0) {
      if (deleteZoomRecording(meetingId, accessToken)) {
        Logger.log(`Deleted recording ${meetingTopic} from Zoom`);
      } 
    }
    
    // Log the transfer to the tracking sheet
    logTransferToSheet(
      trackingSheet, 
      meetingId, 
      meetingTopic, 
      meetingDate, 
      filesTransferred.join(", "), 
      folder.getId(), 
      new Date(),
      hostEmail
    );
    
    Logger.log(`Completed processing recording ${meetingTopic}`);
  } catch (error) {
    Logger.log(`Error processing recording: ${error.message}`);
    trackError(error);
  }
}

/**
 * Process a meeting summary: download, upload to Drive
 */
function processMeetingSummary(meeting, accessToken, trackingSheet) {
  try {
    const meetingId = meeting.meeting_uuid;
    let meetingTopic = meeting.meeting_topic || "Unnamed Meeting";
    
    // Remove topic prefix if configured
    if (CONFIG.TOPIC_PREFIX_TO_REMOVE && meetingTopic.startsWith(CONFIG.TOPIC_PREFIX_TO_REMOVE)) {
      const prefixPattern = new RegExp(`^${CONFIG.TOPIC_PREFIX_TO_REMOVE}(\\s\\|\\s|\\s)`);
      meetingTopic = meetingTopic.replace(prefixPattern, "");
    }
    
    const meetingDate = meeting.summary_created_time ? meeting.summary_created_time.split("T")[0] : "Unknown Date";
    const hostEmail = meeting.meeting_host_email || "Unknown Host";
    
    Logger.log(`Processing meeting summary: ${meetingTopic} (${meetingId}) - Host: ${hostEmail}`);
    
    // Check if we've already processed this meeting
    if (isSummaryProcessed(meetingTopic, meetingDate, trackingSheet)) {
      Logger.log(`Meeting summary for ${meetingTopic} (${meetingId}) already processed. Skipping.`);
      return;
    }
    
    // Get summary data
    const summaryData = getMeetingSummaryDetail(meetingId, accessToken);
    if (!summaryData || !summaryData.detailedSummary) {
      Logger.log(`No summary data available for meeting ${meetingTopic}`);
      return;
    }
    
    // Create or find the folder in Google Drive
    const folderName = `${meetingDate}_${meetingTopic}`;
    const destinationFolderId = getDestinationFolderId(meetingTopic);
    
    let folder;
    try {
      // Try to find existing folder first
      const parentFolder = DriveApp.getFolderById(destinationFolderId);
      const folderIterator = parentFolder.getFoldersByName(folderName);
      
      if (folderIterator.hasNext()) {
        folder = folderIterator.next();
        Logger.log(`Found existing folder: ${folderName}`);
      } else {
        // Create new folder if it doesn't exist
        folder = createDriveFolder(folderName, destinationFolderId);
        Logger.log(`Created new folder: ${folderName}`);
      }
    } catch (error) {
      Logger.log(`Error creating/finding folder: ${error.message}`);
      folder = createDriveFolder(folderName, destinationFolderId);
    }
    
    let filesTransferred = [];

    // Create the summary file
    const summaryFileName = `${meetingDate}_${meetingTopic}_aiSummary.txt`;
    
    // Format summary content
    let summaryContent = "";
    
    // Add meeting metadata
    summaryContent += `#####  *AI Generated* Meeting Summary for ${meetingTopic}  #####\n`;
    summaryContent += `${meetingDate || ""}\n\n`;

    // Add summary overview
    if (summaryData.overviewSummary) {
        summaryContent += `== SUMMARY OVERVIEW ==\n${summaryData.overviewSummary}\n\n`;
    }

    // Add main summary (as labeled sections)
    summaryContent += `== FULL SUMMARY ==\n\n`;
    if (summaryData.detailedSummary && Array.isArray(summaryData.detailedSummary)) {
        summaryData.detailedSummary.forEach(section => {
        // for each { label, summary }
        summaryContent += `-- ${section.label.toUpperCase()} --\n`;
        summaryContent += `${section.summary}\n\n`;
        });
    }
  
    // Add disclaimer
    summaryContent += `=== DISCLAIMER ===\nAI-generated content may be inaccurate or misleading. Always check for accuracy.\n`;
    
    // Create file in Drive
    const file = folder.createFile(summaryFileName, summaryContent, MimeType.PLAIN_TEXT);
    filesTransferred.push(`${summaryFileName} (${file.getId()})`);
    
    // Log the transfer to the tracking sheet
    logTransferToSheet(
      trackingSheet, 
      meetingId, 
      meetingTopic, 
      meetingDate, 
      `AI Summary`, 
      folder.getId(), 
      new Date(),
      hostEmail
    );
    
    Logger.log(`Successfully saved AI summary for meeting ${meetingTopic}`);
    return file;
    
  } catch (error) {
    Logger.log(`Error processing meeting summary: ${error.message}`);
    trackError(error);
    return null;
  }
}

/**
 * Check if a recording has already been processed by looking it up in the tracking sheet
 * Now using meeting topic and date as unique identifiers since Meeting ID column is removed
 */
function isRecordingProcessed(meetingTopic, meetingDate, sheet) {
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === meetingTopic && data[i][1] === meetingDate) {
      return true;
    }
  }
  
  return false;
}

function isSummaryProcessed(meetingTopic, meetingDate, sheet) {
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const topic = data[i][0];
    let date = data[i][1];
    const filesTransferred = data[i][2];

    if (date instanceof Date) {
      date = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    if (topic === meetingTopic && date === meetingDate) {
      if (filesTransferred === "AI Summary") {
        return true;
      } else {
      }
    }
  }
  
  console.log(`[isSummaryProcessed] No matching summary found - returning false`);
  return false;
}

/**
 * Create a folder in Google Drive
 * — throws if parentFolderId is missing or invalid
 */
function createDriveFolder(folderName, parentFolderId) {
  if (!parentFolderId) {
    throw new Error(`createDriveFolder: missing parentFolderId for "${folderName}"`);
  }
  
  let parentFolder;
  try {
    parentFolder = DriveApp.getFolderById(parentFolderId);
  } catch (e) {
    // log the bad ID and rethrow with context
    Logger.log(`⚠️ Invalid parentFolderId "${parentFolderId}" for folder "${folderName}" — ${e.message}`);
    throw new Error(`createDriveFolder: unable to open parentFolderId "${parentFolderId}"`);
  }
  
  // Check if folder already exists
  const folderIterator = parentFolder.getFoldersByName(folderName);
  if (folderIterator.hasNext()) {
    return folderIterator.next();
  }
  
  // Create new folder
  return parentFolder.createFolder(folderName);
}

/**
 * Get the destination folder ID based on the meeting topic
 */
function getDestinationFolderId(meetingTopic) {
  const keywords = Object.keys(CONFIG.FOLDER_MAPPING);
  const topicLower = meetingTopic.toLowerCase();

  for (const keyword of keywords) {
    if (topicLower.includes(keyword.toLowerCase())) {
      return CONFIG.FOLDER_MAPPING[keyword];
    }
  }

  // If no keyword matches, return the default folder ID
  return CONFIG.DEFAULT_FOLDER_ID;
}

/**
 * Download a file from URL and upload it to Google Drive
 */
function downloadAndUploadFile(url, fileName, folderId, mimeType) {
  // Download the file
  const options = {
    method: "get",
    muteHttpExceptions: true,
    followRedirects: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`Failed to download file: ${response.getResponseCode()}`);
    }
    
    const fileBlob = response.getBlob().setName(fileName);
    
    // Set the MIME type on the blob before creating the file
    if (mimeType) {
      fileBlob.setContentType(mimeType);
    }
    
    // Upload to Google Drive
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(fileBlob);
    
    Logger.log(`Successfully transferred file: ${fileName}`);
    return file;
  } catch (error) {
    Logger.log(`Error in downloadAndUploadFile for ${fileName}: ${error.message}`);
    throw error;
  }
}

/**
 * Get MIME type based on file extension
 */
function getMimeType(extension) {
  const mimeTypes = {
    "mp4": "video/mp4",
    "m4a": "audio/m4a",
    "txt": "text/plain",
    "vtt": "text/vtt",
    "json": "application/json",
    "pdf": "application/pdf"
  };
  
  return mimeTypes[extension.toLowerCase()] || null;
}

/**
 * Delete a recording from Zoom
 */
function deleteZoomRecording(meetingId, accessToken) {
  const apiUrl = `https://api.zoom.us/v2/meetings/${meetingId}/recordings`;
  const options = {
    method: "delete",
    headers: {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    return response.getResponseCode() === 200 || response.getResponseCode() === 204;
  } catch (error) {
    Logger.log(`Error deleting Zoom recording: ${error.message}`);
    return false;
  }
}

/**
 * Log a transfer to the tracking sheet
 */
function logTransferToSheet(sheet, meetingId, meetingTopic, meetingDate, filesTransferred, folderId, transferDate, hostEmail) {
  // Get the folder name
  const folder = DriveApp.getFolderById(folderId);
  const folderName = folder.getName();
  const folderUrl = `https://drive.google.com/drive/folders/${folderId}`;
  
  // Extract just the file types from the filesTransferred string
  let fileTypes = [];
  
  if (filesTransferred.includes(".mp4")) {
    fileTypes.push("Video");
  }
  if (filesTransferred.includes(".m4a")) {
    fileTypes.push("Audio");
  }
  if (filesTransferred.includes("chat.txt") || filesTransferred.includes("Chat.txt")) {
    fileTypes.push("Chat");
  }
  if (filesTransferred.includes("transcript.txt")) {
    fileTypes.push("Transcript");
  }
  if (filesTransferred.includes("AI Summary")) {
    fileTypes.push("AI Summary");
  }
  
  const fileTypesList = fileTypes.join(", ");
  
  // Format the meeting date
  let formattedMeetingDate = meetingDate;
  if (meetingDate && meetingDate !== "Unknown Date") {
    try {
      const dateParts = meetingDate.split("-");
      if (dateParts.length === 3) {
        // Assuming ISO format YYYY-MM-DD
        const dateObj = new Date(dateParts[0], parseInt(dateParts[1])-1, dateParts[2]);
        formattedMeetingDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MM-dd-yyyy");
      }
    } catch (e) {
      Logger.log(`Error formatting meeting date: ${e.message}`);
    }
  }
  
  // Add new row to sheet with updated schema
  const rowData = [
    meetingTopic,
    formattedMeetingDate,
    fileTypesList,
    folderName,
    transferDate,
    hostEmail
  ];
  
  // Append the row
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
  
  // Set the folder name as rich text with a link in the "Location in Drive" column (column D, index 4)
  const richText = SpreadsheetApp.newRichTextValue()
    .setText(folderName)
    .setLinkUrl(0, folderName.length, folderUrl)
    .build();
  sheet.getRange(lastRow + 1, 4).setRichTextValue(richText);
}

/**
 * Track errors in a separate sheet
 */
function trackError(error) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.TRACKING_SHEET_ID);
    let errorSheet = spreadsheet.getSheetByName("Errors");
    console.error(`Error in Zoom to Drive connector: ${error.message}`);
    
    if (!errorSheet) {
      errorSheet = spreadsheet.insertSheet("Errors");
      errorSheet.getRange("A1:C1").setValues([["Timestamp", "Error Message", "Stack Trace"]]);
      errorSheet.getRange("A1:C1").setFontWeight("bold");
      errorSheet.setFrozenRows(1);
    }
    
    errorSheet.appendRow([
      new Date(),
      error.message,
      error.stack || "No stack trace available"
    ]);
  } catch (e) {
    Logger.log(`Failed to log error: ${e.message}`);
  }
}

/**
 * Set up a time-based trigger to run the connector daily at 8:00 UTC
 */
function createDailyTrigger() {
  // Delete any existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "runZoomToDriveConnector") {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // Create a new trigger to run daily at 8:00 UTC
  ScriptApp.newTrigger("runZoomToDriveConnector")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .inTimezone("UTC")
    .create();
    
  Logger.log("Created daily trigger to run at 8:00 UTC");
}

/**
 * Setup function - run this manually once to set up the script
 */
function setup() {
  // Initialize tracking sheet
  const trackingSheet = initTrackingSheet();
  
  // Create trigger to run daily
  createDailyTrigger();
  
  // Log setup completion
  Logger.log("Setup complete! The connector will run daily.");
  Logger.log(`Tracking spreadsheet ID: ${CONFIG.TRACKING_SHEET_ID}`);
}