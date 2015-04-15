
var YOUR_EMAIL_ADDRESS = "luxxart@gmail.com"

var ONEDAY = 24 * 60 * 60 * 1000; // hours * minutes * seconds * milliseconds
var totalDaysAgo = 0;  //  no need to be reset after the short loop

/*
	To use tables go to:
	https://developers.google.com/apps-script/advanced/fusion-tables
	https://developers.google.com/fusiontables/docs/v2/getting_started#invoking
	https://developers.google.com/fusiontables/docs/v2/using
	https://drive.google.com/?ddrp=1#query?view=2&filter=tables
    
    message API:
    https://developers.google.com/apps-script/reference/gmail/gmail-message
	
	LOOK AT THE END OF THIS FILE
*/

/*
    TODO:
    Change the days ago calculation formula to take into account the latest message, not the first ("last").
    Charts will change drastically.
*/


function check_ToAnswer_Emails()
{
	createLabelledEmailsReport("___To do/To answer", "[ TAEmAl ]");
}

function check_FollowUp_Emails()
{
	createLabelledEmailsReport("___Follow Up", "[ FUEmAl ]");
}


function createLabelledEmailsReport(LABEL_NAME, MESSAGE_TAG)
{
	// var allMessageInThread = "";
	
	// Log the subject lines of the threads labeled with MyLabel 
	var NUMBER_OF_FIRST_MESSAGES = 22;
	var ss = SpreadsheetApp.getActiveSpreadsheet();

	var threadDate = 0, todaysDate = 0, lastMessage_daysAgo = 0;
	var firstMessageSubject = "";
	var formattedDate = "";
	var sender_HTML = "(&nbsp;";
	var allThreadMessages = "";

	var messageBodyHTML = "<br><br>The latest " + NUMBER_OF_FIRST_MESSAGES + " messages:<br><br><ol>";

	//  Get label reference
	var label = GmailApp.getUserLabelByName(LABEL_NAME);

	//  Get messages, members of that label
	var threads = label.getThreads();
	
	//  Test the function to query Fusion Tables
	// messageBodyHTML += "<strong>Available tables:</strong><br>" + listTables() + "<br>" + "<br>" + "<br>";
	// messageBodyHTML += runQuery("1uAaCQdBgiiWcOG-a1yiR2NomKz_i_S7eJ2X8F7iX", threads.length) + "<br>" + "<br>" + "<br>";
	

	/*
		var Emails2AnswerSheet = ss.getSheetByName("Emails to answer");  
		var range_em2Ans = Emails2AnswerSheet.getRange("a2:a4");
		var myThreads = [1,2,3]
		range_em2Ans.setValues(threads);
	*/	

	//  Create list of the first NUMBER_OF_FIRST_MESSAGES messages
	//  Loop over the messages
	for (var i = 0; i < NUMBER_OF_FIRST_MESSAGES; i++)
	{
		/*
			Logger.log(threads[i].getFirstMessageSubject());
			allMessageInThread = threads[i].getMessages();
			allMessageInThread[0].getDate();
			lastMessage_daysAgo = daydiff(threads[i].getLastMessageDate(),today());
		*/
		formattedDate = Utilities.formatDate(new Date(threads[i].getLastMessageDate()), "GMT", "yyyy-MM-dd");	
		lastMessage_daysAgo = calculateDaysDifference(new Date(threads[i].getLastMessageDate()),new Date(), ONEDAY, 0);

		// get first message subject
		firstMessageSubject = threads[i].getFirstMessageSubject();
		if(firstMessageSubject === "")
			firstMessageSubject = "[ NO SUBJECT ]";


		//  get (and format) the sender of the message

		allThreadMessages = threads[i].getMessages();
        var last_message_number = allThreadMessages.length - 1;
        
		var first_sender = allThreadMessages[0].getFrom();
        var first_recipient = allThreadMessages[0].getTo();        
        var last_sender = allThreadMessages[last_message_number].getFrom();
        var last_recipient = allThreadMessages[last_message_number].getTo();
        var latestMessage_daysAgo = calculateDaysDifference(new Date(allThreadMessages[last_message_number].getDate()),new Date(), ONEDAY, 0);
        
		
		if(first_sender === YOUR_EMAIL_ADDRESS    || first_sender === '"luxxart@gmail.com" <luxxart@gmail.com>'    || first_sender === "Lukasz Przewlocki"    || first_sender === "Lukasz Przewlocki <luxxart@gmail.com>")    first_sender = "[ me ]";
		if(first_recipient === YOUR_EMAIL_ADDRESS || first_recipient === '"luxxart@gmail.com" <luxxart@gmail.com>' || first_recipient === "Lukasz Przewlocki" || first_recipient === "Lukasz Przewlocki <luxxart@gmail.com>") first_recipient = "[ me ]";
        if(last_sender === YOUR_EMAIL_ADDRESS     || last_sender === '"luxxart@gmail.com" <luxxart@gmail.com>'     || last_sender === "Lukasz Przewlocki"     || last_sender === "Lukasz Przewlocki <luxxart@gmail.com>")     last_sender = "[ me ]";
        if(last_recipient === YOUR_EMAIL_ADDRESS  || last_recipient === '"luxxart@gmail.com" <luxxart@gmail.com>'  || last_recipient === "Lukasz Przewlocki"  || last_recipient === "Lukasz Przewlocki <luxxart@gmail.com>")  last_recipient = "[ me ]";
        
        first_sender =    "<strong><span style='background-color:#ff0;'>&nbsp;" + first_sender +    "&nbsp;</span></strong>";
        last_sender =     "<strong><span style='background-color:#ff0;'>&nbsp;" + last_sender +     "&nbsp;</span></strong>";
        first_recipient = "<strong><span style='background-color:#ff0;'>&nbsp;" + first_recipient + "&nbsp;</span></strong>";
        last_recipient =  "<strong><span style='background-color:#ff0;'>&nbsp;" + last_recipient +  "&nbsp;</span></strong>";

        sender_HTML = "(&nbsp;";
        
		if(allThreadMessages.length>1)
		{            
            
			// if(allThreadMessages[1] !== undefined) 
            sender_HTML += last_sender + " --> " + last_recipient;
            sender_HTML += " = the latest of <strong>" + allThreadMessages.length + "</strong>."; // else sender_HTML += "_message[1] UNDEFINED; allThreadMessages.length: " + allThreadMessages.length;
            sender_HTML += " Initially: ";
		}

		sender_HTML += first_sender + " --> " + first_recipient; // // if(allThreadMessages[0] !== undefined) __ else	sender_HTML = "_message[0] UNDEFINED";
        
        
		sender_HTML += "&nbsp;)";

		
		//  Message ID to be used in message URL	
		var link = threads[i].getId();

		//  Alternating background colour
		if (i%2 === 0)
			var colour = "#f0f0f0";
		else
			var colour = "#f9f9f9";

		messageBodyHTML += "<li style='background-color:" + colour + ";'>"; // 'padding:15px; margin-bottom: 15px;
		messageBodyHTML += "<strong> " + latestMessage_daysAgo + "</strong> <span style='color:#964c16;'>(" + lastMessage_daysAgo + ")</span> days ";
		messageBodyHTML += "/ <strong>" + calculateYearsDifference(latestMessage_daysAgo) + "</strong> <span style='color:#964c16;'>(" + calculateYearsDifference(lastMessage_daysAgo) + ")</span> years ago ";
		messageBodyHTML += "<span style='background-color:#FFBFBF;'>&nbsp;[ " + formattedDate + " ] </span>"; // <br>
		//  Display sender
		messageBodyHTML += "<span style='padding-left:44px;'>" + sender_HTML + "</span>:<br>"; // <br>
		//  Display subject
		messageBodyHTML += "<span style='padding-left:88px;'><span style='color:#FFFFFF !important; background-color:#cfc;'>&nbsp;"; 
		messageBodyHTML += "<a href='https://mail.google.com/mail/#all/" + link + "'>"; 
		messageBodyHTML += "<span>" + firstMessageSubject + "</span></a>&nbsp;</span></span> "; //&raquo;
		messageBodyHTML += "</li>";

		/*
			messageBodyHTML += "<a href = '";
			messageBodyHTML += threads[i].getPermalink() + "'>";
			messageBodyHTML += firstMessageSubject + "</a>, ";
		*/
	}
	messageBodyHTML += "</ol>";

	sender_HTML = "";
	// totalDaysAgo = 0;  // no need, because there is a parameter telling not to count when not needed

	//  List of all messages

	messageBodyHTML += "<br><br>All messages:<br><br><ol>";
	for (var i = threads.length - 1; i >= 0; i--)
	{
		/*
			Logger.log(threads[i].getFirstMessageSubject());
			allMessageInThread = threads[i].getMessages();
			allMessageInThread[0].getDate();
			lastMessage_daysAgo = daydiff(threads[i].getLastMessageDate(),today());
		*/
		formattedDate = Utilities.formatDate(new Date(threads[i].getLastMessageDate()), "GMT", "yyyy-MM-dd");

		lastMessage_daysAgo = calculateDaysDifference(new Date(threads[i].getLastMessageDate()),new Date(), ONEDAY, 1);  

		firstMessageSubject = threads[i].getFirstMessageSubject();
		if(firstMessageSubject === "")
			firstMessageSubject = "[ NO SUBJECT ]";

		/*
			//  get the sender of the message	

			allThreadMessages = threads[i].getMessages();
			if(allThreadMessages[0] !== undefined)
			{
			sender_HTML = "[__" + allThreadMessages[0].getFrom() + "";
			}
			else sender_HTML = "_message[0] UNDEFINED";
			
			if(allThreadMessages.length>1)
			{
			if(allThreadMessages[1] !== undefined)
			{
				sender_HTML += "- - - -> " + first_sender + "";
			}
			else sender_HTML += "_message[1] UNDEFINED; allThreadMessages.length: " + allThreadMessages.length;
			}
			sender_HTML += "__]  -  ";
		*/
		
		messageBodyHTML += "<li><strong> " + lastMessage_daysAgo + "</strong> days (<strong>";
		messageBodyHTML += calculateYearsDifference(lastMessage_daysAgo) + "</strong> years) ago - [ ";
		messageBodyHTML += formattedDate + " ] - ";
		messageBodyHTML += sender_HTML + "<a href = '";
		messageBodyHTML += threads[i].getPermalink() + "'>";
		messageBodyHTML += firstMessageSubject + "</a></li>";
	}
	messageBodyHTML += "</ol>";
	
	
	//  list all labels
	
	messageBodyHTML += "<br><br>All labels:<br><br><ol>";
	var labels = GmailApp.getUserLabels();
	for (var i = 0; i < labels.length; i++)
	{
		messageBodyHTML += "<li>" + labels[i].getName() + "</li>";
	}
	messageBodyHTML += "</ol>";

	messageBodyHTML = "Total: <strong>" 
      + numberWithCommas(totalDaysAgo) 
      + "</strong> days ago<br><br>" 
      + "The stats: " + runQuery("1uAaCQdBgiiWcOG-a1yiR2NomKz_i_S7eJ2X8F7iX", threads.length, LABEL_NAME) 
      + "<br>" 
      + "The script: <a href='https://script.google.com/d/169sTxQiDiqvKECPEelbt7WDjuw45AAKptb_yN36vgBQll8W5HoH12Qa5/edit?usp=drive_web'>[ SCR ] - Unactioned emails (email) reminder Script</a><br>"
      + messageBodyHTML;
	

	sendEmailReport(
		"[ " + LABEL_NAME + " ] " + threads.length + " msgs, "
		+ numberWithCommas(totalDaysAgo) + " days ("
		+ Math.floor(totalDaysAgo / 365 * 1) / 1 + " yrs), avg: "
		+ Math.round(totalDaysAgo / threads.length * 10) / 10 + " days ("
		+ Math.round(totalDaysAgo / threads.length / 365 * 100) / 100 + " yrs) "
		+ MESSAGE_TAG, 
		messageBodyHTML, messageBodyHTML);
}

function calculateYearsDifference(daysAgo)
{
	var yDiff = Math.floor(daysAgo / 365 * 10) / 10;
	if (yDiff === Math.floor(daysAgo / 365)) 
		yDiff += ".0";
	return yDiff;
}

function calculateDaysDifference(threadDate,todaysDate, ONEDAY, includeInTotal)
{
	var daysDiff = (todaysDate.getTime() - threadDate.getTime()) / ONEDAY;
	if(includeInTotal === 1) totalDaysAgo += daysDiff;
	return Math.floor(daysDiff);
}

function sendEmailReport(subject, messagePlain, messageBodyHTML) {
	// var message = "This email was sent at " + new Date();
	var recipient = YOUR_EMAIL_ADDRESS;

	MailApp.sendEmail(recipient, subject, messagePlain, { htmlBody: messageBodyHTML });
}

function numberWithCommas(x) {
	// rounding (at the moment nothing after decimal point
	x = Math.floor(x * 1) / 1;
	return x.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}


/**
 * This sample lists Fusion Tables that the user has access to.
 */
function listTables() {
  var tables = FusionTables.Table.list();
  var result_HTML = "";
  var table = "";
  if (tables.items) {
    for (var i = 0; i < tables.items.length; i++) {
      table = tables.items[i];
      result_HTML += " <strong>Table ID:</strong>&nbsp;&nbsp;&nbsp;" + table.tableId 
        + " &nbsp;&nbsp;&nbsp;<strong>Table name:</strong>&nbsp;&nbsp;&nbsp;" + table.name + "<br>";
    }
  } else {
    Logger.log('No tables found.');
  }
  return result_HTML;
}

/**
 * This sample queries for the first 100 rows in the given Fusion Table and
 * saves the results to a new spreadsheet.
 */
function runQuery(tableId, message_numbers, label_name) {
  var sql = 'SELECT * FROM ' + tableId + ' LIMIT 100';
  var result = FusionTables.Query.sqlGet(sql, {
    hdrs: false
  });
  if (result.rows) {
    // var spreadsheet = SpreadsheetApp.create('[ TEST ] - Fusion Table Query Results');
	var spreadsheet = SpreadsheetApp.openById("1Cf2xFK5xOL_pADFbb5u4FprQ0rzq7dmDmpsSxuHDEOI");
    var sheet = spreadsheet.getActiveSheet();

    // Append the headers.
    // sheet.appendRow(result.columns);

    // Append the message numbers.
    if(label_name === "___To do/To answer")
        sheet.appendRow([new Date(), label_name, message_numbers, numberWithCommas(totalDaysAgo)]);
    else
        sheet.appendRow([new Date(), label_name, "", "", message_numbers, numberWithCommas(totalDaysAgo)]);

    // Append the results (list the table entries).
    // sheet.getRange(2, 1, result.rows.length, result.columns.length).setValues(result.rows);

    return 'Query results spreadsheet updated: %s',  spreadsheet.getUrl();
  } else {
    Logger.log('No rows returned.');
  }
}