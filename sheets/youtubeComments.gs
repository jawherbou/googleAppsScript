/**
 * Fetches YouTube comments for a given video ID and inserts them into Google Sheets.
 * The video ID is extracted from the URL provided in cell B1.
 * The comments are inserted starting from row 5, with headers in row 5.
 * Make sure to enable the YouTube Data API v3 in your Google Cloud project and replace
 * the apiKey variable with your actual API key.
 */

function getYouTubeComments() {
  const apiKey = 'INSERT_API_KEY'; // Replace with your API key
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Read url from B1 and extract videoId
  const url = sheet.getRange('B1').getValue().trim(); // Remove any whitespace and ensure no extra characters
  let videoId = url.match(/[?&]v=([^&]+)/)[1];

  const maxResults = 100; // Number of comments to fetch in one API call (max 100)
  let nextPageToken = ''; // Used for pagination
  let comments = [];

  // Clear content from row 5 downwards
  sheet.getRange('4:5000').clearContent(); // Adjust the range as needed

  do {
    // Construct the API URL
    let url = `https://www.googleapis.com/youtube/v3/commentThreads?part=snippet&videoId=${videoId}&maxResults=${maxResults}&key=${apiKey}&pageToken=${nextPageToken}`;

    // Fetch the data from YouTube API
    let response = UrlFetchApp.fetch(url);
    let result = JSON.parse(response.getContentText());

    // Extract comments and push to the array
    result.items.forEach(item => {
      let comment = item.snippet.topLevelComment.snippet.textDisplay;
      let author = item.snippet.topLevelComment.snippet.authorDisplayName;
      let publishedAt = item.snippet.topLevelComment.snippet.publishedAt;
      comments.push([author, comment, publishedAt]);
    });

    // Set nextPageToken for pagination
    nextPageToken = result.nextPageToken;

  } while (nextPageToken);

  // Log the comments or insert into Google Sheets
  Logger.log(comments);
  insertCommentsToSheet(comments);
}

// Helper function to insert comments into Google Sheets
function insertCommentsToSheet(comments) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Insert header at row 5 and apply formatting
  const headerRange = sheet.getRange(5, 1, 1, 3);
  headerRange.setValues([['Author', 'Comment', 'Published At']]);
  headerRange.setFontWeight('bold'); // Make headers bold
  headerRange.setBackground('#f0f0f0'); // Light grey background for headers
  headerRange.setHorizontalAlignment('center'); // Center align headers

  // Insert comments starting from row 5
  if (comments.length > 0) {
    sheet.getRange(6, 1, comments.length, comments[0].length).setValues(comments);
  }
}
