// MAIN ROOT CONFIGURATION - YOUR DRIVE ROOT TARGET ID
var MAIN_ROOT_FOLDER_ID = "1_cFVDveVdMGL-liqJlV6E1AZIuLbhh5S";
var SHEET_ID = "19EtAUIpbN_rqitpJdPv7Hd1xooPXFVTAnVhC17couyk";

// FIX: derive a real mime type from the file extension instead of forcing
// every upload to "application/octet-stream". The video especially needs
// a correct video/mp4 type for reliable playback once served from Drive.
function getMimeType(fileName) {
  var ext = fileName.split('.').pop().toLowerCase();
  var map = {
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'png': 'image/png',
    'mp4': 'video/mp4',
    'mind': 'application/octet-stream'
  };
  return map[ext] || 'application/octet-stream';
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (!data.orderId) {
      throw new Error("Missing orderId in payload");
    }

    // 1. Target the main directory file structure
    var rootFolder = DriveApp.getFolderById(MAIN_ROOT_FOLDER_ID);

    // 2. Automate creating an isolated subfolder named after the unique Order ID
    var subFolder = rootFolder.createFolder(data.orderId);
    subFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // 3. Base64 file upload handler
    function uploadFileToSubfolder(base64Data, fileName) {
      if (!base64Data || !fileName) return "";
      var decoded = Utilities.base64Decode(base64Data);
      var blob = Utilities.newBlob(decoded, getMimeType(fileName), fileName);
      var file = subFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      // FIX: drive.google.com (not docs.google.com) is the current, reliable
      // direct-download host for this link pattern.
      return "https://drive.google.com/uc?export=download&id=" + file.getId();
    }

    var photoUrl = uploadFileToSubfolder(data.photoData, data.photoName);
    var mindUrl = uploadFileToSubfolder(data.mindData, data.mindName);
    var videoUrl = uploadFileToSubfolder(data.videoData, data.videoName);

    // 4. Save metadata to the Sheet
    var sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    var timestamp = new Date();

    sheet.appendRow([
      timestamp,
      data.customerName,
      data.phoneNumber,
      data.orderId,
      photoUrl,
      mindUrl,
      videoUrl
    ]);

    return ContentService.createTextOutput(JSON.stringify({
      "status": "success",
      "orderId": data.orderId,
      "photoUrl": photoUrl,
      "mindUrl": mindUrl,
      "videoUrl": videoUrl
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // This now actually reaches the browser, since the client no longer
    // uses no-cors and can read this response.
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error",
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var orderId = e.parameter.orderId;
    if (!orderId) {
      return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Missing orderId"}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][3] && data[i][3].toString() == orderId) {
        var result = {
          "status": "success",
          "customerName": data[i][1],
          "mindUrl": data[i][5],
          "videoUrl": data[i][6]
        };
        return ContentService.createTextOutput(JSON.stringify(result))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    return ContentService.createTextOutput(JSON.stringify({"status": "not_found"}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
