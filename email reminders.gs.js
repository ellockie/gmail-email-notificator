
// get the email address of the person running the script
var YOUR_EMAIL_ADDRESS = Session.getActiveUser().getEmail();

var ONEDAY = 24 * 60 * 60 * 1000; // hours * minutes * seconds * milliseconds
var totalDaysAgo = 0;  //  no need to be reset after the short loop

var PERSONAL_FOLLOW_UP_LABEL_NAME = "___Follow Up";
var PFU_REPORT_LABEL_NAME = "[ FUEmAl ]";
var TO_ANSWER_LABEL_NAME = "___To do/To answer";
var TA_REPORT_LABEL_NAME = "[ TAEmAl ]";
var WORK_FOLLOW_UP_LABEL_NAME = "___Work/__Follow Up";
var WFU_REPORT_LABEL_NAME = "WFUEmAl";

var init_time = 0;

/*
	To use tables go to:
	 - https://developers.google.com/apps-script/advanced/fusion-tables
	 - https://developers.google.com/fusiontables/docs/v2/getting_started#invoking
	 - https://developers.google.com/fusiontables/docs/v2/using
	 - https://drive.google.com/?ddrp=1#query?view=2&filter=tables

	message API:
	 - https://developers.google.com/apps-script/reference/gmail/gmail-message
    
    javascript code conventions:
     - http://javascript.crockford.com/code.html
     - http://www.w3schools.com/js/js_conventions.asp

	LOOK AT THE END OF THIS FILE FOR FUSION TABLES FUNCTIONS
*/

/*
	TODO:
	 - Change the days ago calculation formula to take into account the latest message, not the first ("last") (numbers and charts will change drastically).
     - If there are problems, start using DB (Fusion Tables)
*/


function generate_all_reports()
{
    report_ToAnswer_emails();
    report_FollowUp_emails();
    report_Work_FollowUp_emails();
}


// Reports for individual labels, if needed
function report_ToAnswer_emails()      { report_emails( TO_ANSWER_LABEL_NAME,          TA_REPORT_LABEL_NAME ); }
function report_FollowUp_emails()      { report_emails( PERSONAL_FOLLOW_UP_LABEL_NAME, PFU_REPORT_LABEL_NAME ); }
function report_Work_FollowUp_emails() { report_emails( WORK_FOLLOW_UP_LABEL_NAME,     WFU_REPORT_LABEL_NAME ); }


function report_emails(THREAD_LABEL, REPORT_LABEL)
{
    init_time = new Date();
    totalDaysAgo = 0;
	create_labelled_emails_report(THREAD_LABEL, REPORT_LABEL);
}


function create_labelled_emails_report(CURRENT_LABEL_NAME, CURRENT_REPORT_LABEL)
{
	// var allMessageInThread = "";

	// Log the subject lines of the threads labeled with MyLabel
	var NUMBER_OF_FIRST_MESSAGES = 22;
	var ss = SpreadsheetApp.getActiveSpreadsheet();

	var threadDate = 0, todaysDate = 0, firstMessage_daysAgo = 0, latestMessage_daysAgo = 0;
	var last_message_number = 0;
	var firstMessageSubject = "";
	var formattedDate = "";
	var sender_HTML = "(&nbsp;";
	var allThreadMessages = "";

	var messageBodyHTML = "<br><br>The latest " + NUMBER_OF_FIRST_MESSAGES + " messages:<br><br><ol>";

	//  Get label reference
	var label = GmailApp.getUserLabelByName(CURRENT_LABEL_NAME);

	//  Get messages, members of that label
	var threads = label.getThreads();

	//  Test the function to query Fusion Tables
	// messageBodyHTML += "<strong>Available tables:</strong><br>" + list_tables() + "<br>" + "<br>" + "<br>";
	// messageBodyHTML += run_query("1uAaCQdBgiiWcOG-a1yiR2NomKz_i_S7eJ2X8F7iX", threads.length) + "<br>" + "<br>" + "<br>";


	/*
		var Emails2AnswerSheet = ss.getSheetByName("Emails to answer");
		var range_em2Ans = Emails2AnswerSheet.getRange("a2:a4");
		var myThreads = [1,2,3]
		range_em2Ans.setValues(threads);
	*/



    /**********************************************************************************/
    //  Loop over the first NUMBER_OF_FIRST_MESSAGES messages
    /**********************************************************************************/

    for (var i = 0; i < NUMBER_OF_FIRST_MESSAGES; i++)
	{
		/*
			Logger.log(threads[i].getFirstMessageSubject());
			allMessageInThread = threads[i].getMessages();
			allMessageInThread[0].getDate();
			firstMessage_daysAgo = daydiff(threads[i].getLastMessageDate(),today());
		*/
		formattedDate = Utilities.formatDate(new Date(threads[i].getLastMessageDate()), "GMT", "yyyy-MM-dd");

		// get first message subject
		firstMessageSubject = threads[i].getFirstMessageSubject();
		if(firstMessageSubject === "")
			firstMessageSubject = "[ NO SUBJECT ]";


		//  get (and format) the sender of the message

		allThreadMessages = threads[i].getMessages();
		last_message_number = allThreadMessages.length - 1;

		var first_sender = allThreadMessages[0].getFrom();
		var first_recipient = allThreadMessages[0].getTo();
		var last_sender = allThreadMessages[last_message_number].getFrom();
		var last_recipient = allThreadMessages[last_message_number].getTo();
		// firstMessage_daysAgo = calculate_days_difference(new Date(threads[i].getLastMessageDate()),new Date());
		firstMessage_daysAgo = calculate_days_difference(new Date(allThreadMessages[0].getDate()),new Date());
		latestMessage_daysAgo = calculate_days_difference(new Date(allThreadMessages[last_message_number].getDate()),new Date());


        var me_string = "<span style='background-color:#e8ac43;'>[ me ]</span>";
		if(first_sender === YOUR_EMAIL_ADDRESS	|| first_sender === '"luxxart@gmail.com" <luxxart@gmail.com>'	|| first_sender === "Lukasz Przewlocki"	|| first_sender === "Lukasz Przewlocki <luxxart@gmail.com>")	      first_sender = me_string;
		if(first_recipient === YOUR_EMAIL_ADDRESS || first_recipient === '"luxxart@gmail.com" <luxxart@gmail.com>' || first_recipient === "Lukasz Przewlocki" || first_recipient === "Lukasz Przewlocki <luxxart@gmail.com>") first_recipient = me_string;
		if(last_sender === YOUR_EMAIL_ADDRESS	 || last_sender === '"luxxart@gmail.com" <luxxart@gmail.com>'	 || last_sender === "Lukasz Przewlocki"	 || last_sender === "Lukasz Przewlocki <luxxart@gmail.com>")	      last_sender = me_string;
		if(last_recipient === YOUR_EMAIL_ADDRESS  || last_recipient === '"luxxart@gmail.com" <luxxart@gmail.com>'  || last_recipient === "Lukasz Przewlocki"  || last_recipient === "Lukasz Przewlocki <luxxart@gmail.com>")  last_recipient = me_string;

		first_sender =	"<strong><span style='background-color:#ff0;'>&nbsp;" + first_sender +	"&nbsp;</span></strong>";
		last_sender =	 "<strong><span style='background-color:#ff0;'>&nbsp;" + last_sender +	 "&nbsp;</span></strong>";
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
		messageBodyHTML += "<strong> " + Math.floor(latestMessage_daysAgo) + "</strong> <span style='color:#964c16;'>(" + Math.floor(firstMessage_daysAgo) + ")</span> days ";
		messageBodyHTML += "/ <strong>" + calculate_years_difference(latestMessage_daysAgo) + "</strong> <span style='color:#964c16;'>(" + calculate_years_difference(firstMessage_daysAgo) + ")</span> years ago ";
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


    /**********************************************************************************/
	//  List of all messages
    /**********************************************************************************/

	messageBodyHTML += "<br><br>All messages:<br><br><ol>";
	for (var i = threads.length - 1; i >= 0; i--)
	{
		/*
			Logger.log(threads[i].getFirstMessageSubject());
			allMessageInThread = threads[i].getMessages();
			allMessageInThread[0].getDate();
			firstMessage_daysAgo = daydiff(threads[i].getLastMessageDate(),today());
		*/

		formattedDate = Utilities.formatDate(new Date(threads[i].getLastMessageDate()), "GMT", "yyyy-MM-dd");
        // Utilities.sleep(111);
        
        // try to read from the Fusion Table and feed JSON
        // http://stackoverflow.com/questions/20881213/converting-json-object-into-javascript-array
        // http://stackoverflow.com/questions/684672/loop-through-javascript-object

        if(CURRENT_LABEL_NAME !== PERSONAL_FOLLOW_UP_LABEL_NAME)
        {
            // that's the proper way
            allThreadMessages = threads[i].getMessages();
            last_message_number = allThreadMessages.length - 1;
            firstMessage_daysAgo = calculate_days_difference(new Date(allThreadMessages[0].getDate()),new Date());
            latestMessage_daysAgo = calculate_days_difference(new Date(allThreadMessages[last_message_number].getDate()),new Date());
        }
        else
        {
            // this is wrong, but currently there are too many calls to threads[i].getMessages()
            firstMessage_daysAgo = calculate_days_difference(new Date(threads[i].getLastMessageDate()),new Date());
            latestMessage_daysAgo = firstMessage_daysAgo;
        }
		
		// count the total delay in days, consider the latest message in each thread only
		totalDaysAgo += latestMessage_daysAgo;

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

		messageBodyHTML += "<li><strong> " + Math.floor(firstMessage_daysAgo) + "</strong> days (<strong>";
		messageBodyHTML += calculate_years_difference(firstMessage_daysAgo) + "</strong> years) ago - [ ";
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
	  + number_with_commas(totalDaysAgo)
	  + "</strong> days ago<br><br>"
	  + "The stats: " + run_query("1uAaCQdBgiiWcOG-a1yiR2NomKz_i_S7eJ2X8F7iX", threads.length, CURRENT_LABEL_NAME)
	  + "<br>"
	  + "The script: <a href='https://script.google.com/d/169sTxQiDiqvKECPEelbt7WDjuw45AAKptb_yN36vgBQll8W5HoH12Qa5/edit?usp=drive_web'>[ SCR ] - Unactioned emails (email) reminder Script</a><br>"
	  + messageBodyHTML;


	send_email_report(
		"[ " + CURRENT_LABEL_NAME + " ] " + threads.length + " msgs, "
		+ number_with_commas(totalDaysAgo) + " days ("
		+ Math.floor(totalDaysAgo / 365 * 1) / 1 + " yrs), avg: "
		+ Math.round(totalDaysAgo / threads.length * 10) / 10 + " days ("
		+ Math.round(totalDaysAgo / threads.length / 365 * 100) / 100 + " yrs) "
		+ CURRENT_REPORT_LABEL,
		messageBodyHTML, messageBodyHTML);
}


function calculate_years_difference(daysAgo)
{
	var yDiff = Math.floor(daysAgo / 365 * 10) / 10;
	if (yDiff === Math.floor(daysAgo / 365))
		yDiff += ".0";
	return yDiff;
}


function calculate_days_difference(threadDate, todaysDate)
{
	return (todaysDate.getTime() - threadDate.getTime()) / ONEDAY;
}


function send_email_report(subject, messagePlain, messageBodyHTML)
{
	// var message = "This email was sent at " + new Date();
	var recipient = YOUR_EMAIL_ADDRESS;

	MailApp.sendEmail(recipient, subject, messagePlain, { htmlBody: messageBodyHTML });
}


function number_with_commas(x)
{
	// rounding (at the moment nothing after decimal point
	x = Math.floor(x * 1) / 1;
	return x.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}


/**
 * This sample lists Fusion Tables that the user has access to.
 */
function list_tables()
{
    var tables = FusionTables.Table.list();
    var result_HTML = "";
    var table = "";
    if (tables.items)
    {
        for (var i = 0; i < tables.items.length; i++)
        {
            table = tables.items[i];
            result_HTML += " <strong>Table ID:</strong>&nbsp;&nbsp;&nbsp;" + table.tableId
              + " &nbsp;&nbsp;&nbsp;<strong>Table name:</strong>&nbsp;&nbsp;&nbsp;" + table.name + "<br>";
        }
    }
    else
    {
        Logger.log('No tables found.');
    }
    return result_HTML;
}


/**
 * This sample queries for the first 100 rows in the given Fusion Table and
 * saves the results to a new spreadsheet.
 */
function run_query(tableId, message_numbers, CURRENT_LABEL_NAME)
{
    var sql = 'SELECT * FROM ' + tableId + ' LIMIT 100';
    var result = FusionTables.Query.sqlGet(sql, { hdrs: false });
    if (result.rows)
    {
        // var spreadsheet = SpreadsheetApp.create('[ TEST ] - Fusion Table Query Results');
        var spreadsheet = SpreadsheetApp.openById("1Cf2xFK5xOL_pADFbb5u4FprQ0rzq7dmDmpsSxuHDEOI");
        var sheet = spreadsheet.getActiveSheet();
    
        // Append the headers.
        // sheet.appendRow(result.columns);
    
        // Append the message numbers.
        if(CURRENT_LABEL_NAME === TO_ANSWER_LABEL_NAME)
            sheet.appendRow([new Date(), CURRENT_LABEL_NAME, message_numbers, number_with_commas(totalDaysAgo), "", "", "", "", Math.floor((new Date() - init_time)/1000)]);
        else if(CURRENT_LABEL_NAME === PERSONAL_FOLLOW_UP_LABEL_NAME)
            sheet.appendRow([new Date(), CURRENT_LABEL_NAME, "", "", message_numbers, number_with_commas(totalDaysAgo), "", "", "", Math.floor((new Date() - init_time)/1000)]);
        else
            sheet.appendRow([new Date(), CURRENT_LABEL_NAME, "", "", "", "", message_numbers, number_with_commas(totalDaysAgo), "", "", Math.floor((new Date() - init_time)/1000)]);
    
        // Append the results (list the table entries).
        // sheet.getRange(2, 1, result.rows.length, result.columns.length).setValues(result.rows);
    
        return 'Query results spreadsheet updated: %s',  spreadsheet.getUrl();
    }
    else 
    {
        Logger.log('No rows returned.');
    }
}