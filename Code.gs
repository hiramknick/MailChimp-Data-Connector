var APIMaster = 'XXXXXXXXXXX-xx0';
var ListMasterID = 'xxxxxxx';



function onOpen()
{
  // Placeholder
  
}

function GetDataFromHistory()
{

  var offset = 0;
  var count = 500;
  
  for (var i = 0; i < 10; i++) {
      if (i == 0)
      {
        writeHeader();
      }
      var chimpdata= [];
      offset = (i)*count;
      chimpdata =  mailchimpReports(offset,count);
      writeToSpreadsheet(chimpdata,offset);
    }
  
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('stats').sort(2);
}

function writeHeader ()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('stats');
  
  sheet.clear(); // Clear MailChimp data in Spreadsheet
  // Put Headers
  sheet.appendRow([
    "ID","Sent Time","Campaign Title", "CampaignID","List Name", "List ID","Subject Line", "Recipients","Unique Opens","Total Opens","Total Clicks","Unique Clicks","Unique Subscriber Clicks","Industry_OpenRate","Industry_Click Rate","Industry_Bounce Rate","Industry_UnOpen Rate","Industry_Unsub Rate","Industry Abuse Rate","Open Rate","Click Rate","Bounce Rate","Unopen Rate","Unsub rate","Abuse Rate","Hard Bounces","Soft Bounces","Total Bounces","Unsubscribed","Abuse Reports"
  ]);
  sheet.setFrozenRows(1);
}
  
function writeToSpreadsheet(data,starting)
{
  // select the campaign output sheet

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('stats');
    var numRows = data.length;
    var numCols = data[0].length;
    sheet.getRange(5 + starting,1,numRows,numCols).setValues(data);
}

// Get all Mailchimp reports
function mailchimpReports(offset,count) {

  var API_KEY = APIMaster;
  var dc = API_KEY.split('-')[1];
  var LIST_ID = ListMasterID;
  ;
  var maxCount = 10;
  var skip = 0;
  if (count == null){
    maxCount = 500;
  }
  else
  {
    maxCount = count;
  }
  if (offset == null)
  {
    skip = 0;
  }
  else
  {
    skip = offset;
  }
  
  //var LastSentTime = '2015-05-20T01:01'
  
  // URL and params for the Mailchimp API
  var root = 'https://'+ dc +'.api.mailchimp.com/3.0';
  var endpoint = '/reports?count=' + maxCount + '&offset=' + skip;
  
  // parameters for url fetch
  var params = {
    'method': 'GET',
    'muteHttpExceptions': true,
    'headers': {
      'Authorization': 'apikey ' + API_KEY
    }
  };
  
  try {
    // call the Mailchimp API
    var response = UrlFetchApp.fetch(root+endpoint, params);
    var data = response.getContentText();
    var json = JSON.parse(data);
    
    // get just campaign data
    var reports = json['reports'];
    
    // blank array to hold the campaign data for Sheet
    var reportData = [];
  

    
    // Add the campaign data to the array
    for (var i = 0; i < reports.length; i++) {
      
      // put the campaign data into a double array for Google Sheets
      if (reports[i]["emails_sent"] != 0) {
        
        //Format Date
        //var dateRaw = reports[i]["send_time"];
        //var dateVal = new Date(dateRaw);
        //var formatted = Utilities.formatDate(new Date(reports[i]["send_time"]), "EDT", "yyyy-MM-dd HH:mm:ss");
        //var dateForm = Utilities.formatDate(JSON.stringify(reports[i]["send_time"]), "EDT", "yyyy-MM-dd HH:mm:ss");
        var test = Utilities.formatDate(new Date(reports[i]["send_time"]), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");

        reportData.push([
          i,
          Utilities.formatDate(new Date(reports[i]["send_time"]), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss"),
          reports[i]["campaign_title"],
          reports[i]["id"],
          reports[i]["list_name"],
          reports[i]["list_id"],
          reports[i]["subject_line"],
          reports[i]["emails_sent"],
          reports[i]["opens"]["unique_opens"],
          reports[i]["opens"]["opens_total"],
          reports[i]["clicks"]["clicks_total"],
          reports[i]["clicks"]["unique_clicks"],
          reports[i]["clicks"]["unique_subscriber_clicks"],
          // Industry Metrics
          formatdecimal(reports[i]["industry_stats"]["open_rate"],5),
          formatdecimal(reports[i]["industry_stats"]["click_rate"],5),
          formatdecimal(reports[i]["industry_stats"]["bounce_rate"],5),
          formatdecimal(reports[i]["industry_stats"]["unopen_rate"],5),
          formatdecimal(reports[i]["industry_stats"]["unsub_rate"],5),
          formatdecimal(reports[i]["industry_stats"]["abuse_rate"],8),
          // Stats
          formatdecimal(reports[i]["opens"]["open_rate"],5),
          formatdecimal(reports[i]["clicks"]["click_rate"],5),
          formatdecimal(((reports[i]["bounces"]["hard_bounces"] + reports[i]["bounces"]["soft_bounces"])/reports[i]["emails_sent"]),5),
          formatdecimal(1-reports[i]["opens"]["open_rate"],5),
          formatdecimal((reports[i]["unsubscribed"]/reports[i]["emails_sent"]),5),
          formatdecimal((reports[i]["abuse_reports"]/reports[i]["emails_sent"]),5),

          //DeliveryFailures
          reports[i]["bounces"]["hard_bounces"],
          reports[i]["bounces"]["soft_bounces"],
          reports[i]["bounces"]["hard_bounces"] + reports[i]["bounces"]["soft_bounces"],

          //Notes
          reports[i]["unsubscribed"],
          reports[i]["abuse_reports"]
        ]);
      }
      else {
        reportData.push([
          i,
          "No DATA"
        ]);
      }
    }

    
    // Log the reportData array
    Logger.log(reportData);
    
    
    
    return reportData
    

  }
  catch (error) {
    // deal with any errors
    Logger.log(error);
  };
  
}
function formatdecimal(x, n){
    x = parseFloat(x);
    n = n || 2;
    return parseFloat(x.toFixed(n));
}

