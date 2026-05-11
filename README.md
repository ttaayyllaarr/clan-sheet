# RuneScape Clan Sheet Sync (Google Apps Script)

This Google Apps Script automates the process of syncing RuneScape clan member data, including ranks, XP, RunePixels IDs, Citadel visit dates, join dates, and last activity, directly into a Google Sheet. It leverages both the official RuneScape clan hiscores and the RunePixels API to provide a comprehensive and up-to-date overview of your clan.

## Features

-   **Automated Data Fetching**: Retrieves the latest clan member list from RuneScape's official hiscores.
-   **RunePixels Integration**: Fetches additional player data (RunePixels ID, join date, last activity) from the RunePixels API.
-   **ID Resolution**: Automatically resolves missing RunePixels IDs for clan members.
-   **Citadel Tracking**: Tracks and highlights members who have missed their weekly Citadel cap.
-   **Activity Heatmap**: Applies conditional formatting to visualize member activity levels (e.g., active today, 7d+ inactive, 14d+ inactive, 21d+ inactive).
-   **Configurable Reset Times**: Allows customization of the clan's weekly reset day and time for accurate Citadel tracking.
-   **Audit Mode**: Supports auditing past weeks' Citadel activity by setting a specific "Check Date".
-   **Error Handling**: Includes basic error handling and toast notifications for user feedback.

## How It Works

The script operates in four main stages, which can be run individually or as a "Full Sync":

1.  **Stage 1: Refresh Member List**: Fetches the current clan member list from RuneScape and initial RunePixels data, populating the sheet with RSNs, Ranks, Clan XP, RunePixels IDs, Join Dates, and Last Activity.
2.  **Stage 2: Resolve Missing IDs**: Attempts to find and populate any missing RunePixels IDs for members.
3.  **Stage 3: Update Data**: Performs a POST request to RunePixels to get the most up-to-date Citadel and activity data, then updates the corresponding columns in the sheet.
4.  **Stage 4: Apply Formatting**: Applies all conditional formatting rules (Citadel missed, activity heatmap) and sets column widths and number formats.

## Setup

1.  **Create a New Google Sheet**: Go to [Google Sheets](https://docs.google.com/spreadsheets/u/0/) and create a new blank spreadsheet.
2.  **Open Apps Script Editor**: In the new Google Sheet, go to `Extensions > Apps Script`. You may encounter issues here if you are signed in to more than one Google account.
3.  **Copy and Paste Code**: Delete any existing code in the `Code.gs` file (or create a new `.js` file if preferred) and paste the entire content of the provided `clansheet.js` file into the editor.
4.  **Save Project**: Click the save icon (floppy disk) or press `Ctrl + S` (Windows) / `Cmd + S` (Mac). You might be prompted to name your project; choose a descriptive name like "RuneScape Clan Sync".
5.  **Authorize Script**: The first time you run any function from the script (e.g., `onOpen` or `RunFullSync`), Google will ask you to authorize it. Follow the prompts:
    *   Click "Review permissions".
    *   Select your Google account.
    *   Click "Allow" to grant the necessary permissions (e.g., to connect to external services, edit your spreadsheets).

## Usage

Once set up, you can use the custom menu "RS Clan Sync" in your Google Sheet:

-   **Full Sync (All Stages)**: Runs all stages sequentially to refresh all data and formatting. This is the recommended option for a complete update.
-   **Stage 1: Refresh Member List**: Updates the basic member list from RuneScape and initial RunePixels data.
-   **Stage 2: Resolve Missing IDs**: Specifically targets and resolves any RunePixels IDs that were not found in Stage 1.
-   **Stage 3: Update Data**: Fetches and updates Citadel visit dates, join dates, and last activity from RunePixels.
-   **Stage 4: Apply Formatting**: Re-applies all conditional formatting and column settings.

## Configuration

-   **B1 (Clan Name)**: Enter the exact name of your RuneScape clan. This is crucial for fetching data.
-   **B2 (Check Date)**: (Optional) For "Audit Mode". Enter a specific date here to check Citadel activity *up to* that date. Clear the cell to return to current week tracking.
-   **E2 (Reset Time (UTC))**: Enter the UTC time when your clan's weekly reset occurs (e.g., `00:00`).
-   **G2 (Reset Day)**: Select the day of the week when your clan's weekly reset occurs.

## API Endpoints Used

-   **RuneScape Official Hiscores**: `<https://secure.runescape.com/m=clan-hiscores/members_lite.ws>`
-   **RunePixels Player ID Lookup**: `<https://api.runepixels.com/players/{rsn}>`
-   **RunePixels Clan List**: `<https://api.runepixels.com/clans/{clanId}/list>` (POST request)
-   **RunePixels Clan Citadel**: `<https://api.runepixels.com/clans/{clanId}/players/citadel>` (GET request)

## Troubleshooting/Notes

-   **Authorization**: Ensure you have authorized the script to run.
-   **API Rate Limits**: Excessive calls to RunePixels APIs might lead to temporary blocking. The script attempts to be efficient but be mindful of frequent manual runs.
-   **Clan Name Accuracy**: The clan name in cell `B1` must exactly match your clan's name in RuneScape.
-   **Private RuneMetrics**: Players with private RuneMetrics profiles on RuneScape may have "N/A" for join dates or activity if RunePixels cannot access that data.
-   **Script Runtime**: Google Apps Script has a maximum execution time limit (typically 6 minutes). For very large clans, `Stage2_ResolveIDs` might need to be run multiple times if it times out.
-   **Encoding**: The script attempts to handle special characters in RSNs, but very unusual characters might still cause issues.

## Credits

Powered by data from RuneScape and RunePixels.com.

## License

This project is open-sourced under the MIT License.
```
