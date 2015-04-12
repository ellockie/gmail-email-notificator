var ONEDAY = 24*60*60*1000; // hours*minutes*seconds*milliseconds
var totalDaysAgo = 0;  //  for testing purposes only?



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

	var threadDate = 0, todaysDate = 0, daysAgo = 0;
	var firstMessageSubject = "";
	var formattedDate = "";
	var sender_HTML = "( ";
	var allThreadMessages = "";

	var messageBodyHTML = "<br><br>The latest NUMBER_OF_FIRST_MESSAGES messages:<br><br><ol>";

	//  Get label reference
	var label = GmailApp.getUserLabelByName(LABEL_NAME);

	//  Get messages, members of that label
	var threads = label.getThreads();

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
		daysAgo = daydiff(threads[i].getLastMessageDate(),today());
	*/
	formattedDate = Utilities.formatDate(new Date(threads[i].getLastMessageDate()), "GMT", "yyyy-MM-dd");	
	daysAgo = calculateDaysDifference(new Date(threads[i].getLastMessageDate()),new Date(), ONEDAY, 0);

	firstMessageSubject = threads[i].getFirstMessageSubject();
	if(firstMessageSubject==="") firstMessageSubject="[ NO SUBJECT ]";


	//  get (and format) the sender of the message

	allThreadMessages = threads[i].getMessages();
	var sender_email = allThreadMessages[0].getFrom();
	if(sender_email === "luxxart@gmail.com") sender_email = "me";

	if(allThreadMessages[0] !== undefined) {
		sender_HTML = "( " + sender_email;
	}
	else sender_HTML ="_message[0] UNDEFINED";

	if(allThreadMessages.length>1) {
		if(allThreadMessages[1] !== undefined) {
			sender_HTML += " --> " + sender_email;
		}
		else sender_HTML +="_message[1] UNDEFINED; allThreadMessages.length: "+allThreadMessages.length;
	}
	sender_HTML += " )";

	//  Message ID to be used in message URL	
	var link = threads[i].getId();

	//  Alternating background colour
	if (i%2 === 0) var colour = "#f0f0f0";
	else var colour = "#f9f9f9";

	messageBodyHTML += "<li style='background-color:" + colour + ";'>"; // 'padding:15px; margin-bottom: 15px;
	messageBodyHTML += "<strong> " + daysAgo + "</strong> days ";
	messageBodyHTML += "/ <strong>" + calculateYearsDifference(daysAgo) + "</strong> years ago - ";
	messageBodyHTML += "<span style='background-color:#FFDFBF;'>&nbsp;[ " + formattedDate + " ] </span>"; // <br>
	//  Display sender
	messageBodyHTML += "<span style='padding-left:22px;'><span style='background-color:#ff0;'>" + sender_HTML + "</span></span>:<br>"; // <br>
	//  Display topic
	messageBodyHTML += "<span style='padding-left:44px;'><span style='color:#FFFFFF !important; background-color:#DFFFBF;'>&nbsp;"; 
	messageBodyHTML += "<a href='https://mail.google.com/mail/#all/" + link + "'>"; 
	messageBodyHTML += firstMessageSubject + " </a></span></span> "; //&raquo;
	messageBodyHTML += "</li>";

	/*
		messageBodyHTML += "<a href = '";
		messageBodyHTML += threads[i].getPermalink()+"'>";
		messageBodyHTML += firstMessageSubject+"</a>, ";
	*/
	}
	messageBodyHTML += "</ol>";

	sender_HTML = "";


	//  List of all messages

	messageBodyHTML += "<br><br>All messages:<br><br><ol>";
	for (var i = threads.length-1; i >= 0; i--)
	{
	//  Logger.log(threads[i].getFirstMessageSubject());
	//  allMessageInThread = threads[i].getMessages();
	//  allMessageInThread[0].getDate();
	formattedDate = Utilities.formatDate(new Date(threads[i].getLastMessageDate()), "GMT", "yyyy-MM-dd");
	// daysAgo = daydiff(threads[i].getLastMessageDate(),today());

	daysAgo = calculateDaysDifference(new Date(threads[i].getLastMessageDate()),new Date(), ONEDAY, 1);  

	firstMessageSubject = threads[i].getFirstMessageSubject();
	if(firstMessageSubject==="") firstMessageSubject="[ NO SUBJECT ]";

	/*
		//  get the sender of the message	

		allThreadMessages = threads[i].getMessages();
		if(allThreadMessages[0] !== undefined)
		{
		sender_HTML = "[__"+allThreadMessages[0].getFrom()+"";
		}
		else sender_HTML ="_message[0] UNDEFINED";
		
		if(allThreadMessages.length>1)
		{
		if(allThreadMessages[1] !== undefined)
		{
			sender_HTML += "- - - -> " + sender_email+"";
		}
		else sender_HTML +="_message[1] UNDEFINED; allThreadMessages.length: "+allThreadMessages.length;
		}
		sender_HTML += "__]  -  ";
	*/
	
	messageBodyHTML += "<li><strong> " + daysAgo + "</strong> days (<strong>";
	messageBodyHTML += calculateYearsDifference(daysAgo) + "</strong> years) ago - [ ";
	messageBodyHTML += formattedDate + " ] - ";
	messageBodyHTML += sender_HTML + "<a href = '";
	messageBodyHTML += threads[i].getPermalink() + "'>";
	messageBodyHTML += firstMessageSubject + "</a></li>";
	}
	messageBodyHTML += "</ol>";
	
	
	//  list all labels
	
	messageBodyHTML += "<br><br>All labels:<br><br><ol>";
	var labels = GmailApp.getUserLabels();
	for (var i = 0; i < labels.length; i++) {
	messageBodyHTML += "<li>"+labels[i].getName()+"</li>";
	}
	messageBodyHTML += "</ol>";

	messageBodyHTML = "Total: <strong>"+numberWithCommas(totalDaysAgo)+ "</strong> days ago<br>"+messageBodyHTML;

	sendEmailReport(
		"[ "+LABEL_NAME+" ] "+threads.length+" msgs, "
		+ numberWithCommas(totalDaysAgo)+" days ("
		+ Math.floor(totalDaysAgo/365*1)/1+" yrs), avg: "
		+ Math.round(totalDaysAgo/threads.length*10)/10+" days ("
		+ Math.round(totalDaysAgo/threads.length/365*100)/100+" yrs) "
		+ MESSAGE_TAG, 
		messageBodyHTML, messageBodyHTML);
}

function calculateYearsDifference(daysAgo)
{
	var yDiff = Math.floor(daysAgo/365*10)/10;
	if (yDiff === Math.floor(daysAgo/365)) yDiff+=".0";
	return yDiff;
}

function calculateDaysDifference(threadDate,todaysDate, ONEDAY, includeInTotal)
{
	var daysDiff = (todaysDate.getTime() - threadDate.getTime())/ONEDAY;
	if(includeInTotal===1) totalDaysAgo += daysDiff;
	return Math.floor(daysDiff);
}

function sendEmailReport(subject, messagePlain, messageBodyHTML) {
	// var message = "This email was sent at " + new Date();
	var recipient = "luxxart@gmail.com";

	MailApp.sendEmail(recipient, subject, messagePlain, { htmlBody: messageBodyHTML });
}

function numberWithCommas(x) {
	x = Math.floor(x*1)/1; // rounding (at the moment nothing after decimal point
	return x.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}