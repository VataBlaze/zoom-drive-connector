# Zoom to Google Drive Connector

This project provides a Google Apps Script that connects Zoom to Google Drive, allowing users to transfer Zoom cloud recordings and AI meeting summaries directly to their Google Drive. The script also tracks these transfers in a Google Sheet for easy management and reference.

## Features

- Authenticate with the Zoom API using Server-to-Server OAuth.
- Fetch cloud recordings and AI meeting summaries from Zoom.
- Transfer recordings and summaries to specified folders in Google Drive.
- Automatically log transfers in a Google Sheet for tracking purposes.
- Customizable folder mapping based on meeting topics.

## Setup Instructions

1. **Clone the Repository**:
   Clone this repository to your local machine using:

   ```bash
   git clone https://github.com/yourusername/zoom-drive-connector.git
   ```

2. **Configure the Script**:
   - Open the `src/template_zoom-drive-connector_github.js` file.
   - Replace all placeholder values in the `CONFIG` object with your actual credentials:
     - `ZOOM_ACCOUNT_ID`
     - `ZOOM_CLIENT_ID`
     - `ZOOM_CLIENT_SECRET`
     - `GOOGLE_DRIVE_FOLDER_ID`
     - `FOLDER_MAPPING` for specific meeting topics and their corresponding Google Drive folder IDs.
     - `DEFAULT_FOLDER_ID` for recordings that don't match any keywords.

3. **Set Up Zoom API**:
   - Create a Server-to-Server OAuth app in the Zoom Marketplace.
   - Obtain your API credentials (Client ID and Client Secret).

4. **Create Google Drive Folders**:
   - Create the necessary folders in Google Drive and note their IDs.

5. **Initialize the Tracking Sheet**:
   - Run the `setup()` function in the script to initialize the tracking sheet and set up a daily trigger.

6. **Run the Connector**:
   - Test the setup by running the `runZoomToDriveConnector()` function manually to transfer recordings and summaries.

## Security Notice

**Never commit your real API credentials (Zoom or Google) to any public repository.**  
Always use placeholder values in shared code and keep your credentials private.

## Deploying the Script

1. Go to [Google Apps Script](https://script.google.com).
2. Create a new project.
3. Copy the contents of `src/template_zoom-drive-connector_github.js` into the script editor.
4. Save and follow the setup instructions above.

## Usage

Once the setup is complete, the script can be run to automatically fetch and transfer Zoom recordings and summaries to Google Drive. The tracking sheet will log each transfer, making it easy to keep track of what has been processed.

## License

This project is licensed under the MIT License. See the LICENSE file for details.
