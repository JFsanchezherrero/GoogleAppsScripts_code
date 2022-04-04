// --------------------------------------------
// Given a string filter retrieved messages
// --------------------------------------------
function getRelevantMessages(filter_str) {
  // filter_str: string with information to use for the filtering of messages
  // example: 
  // in:sent AND newer_than:24h AND -label:gap_parsed
  // from:someone@gmail.com 
  
  var threads = GmailApp.search(filter_str,0,500); // Limit 500 each time
  var messages=[];
  threads.forEach(
    function(thread) {
      //messages.push(thread.getMessages()[0]); 
      // This code only returns first message in thread
      // but I would prefer that it
      // returns all messages in each thread
      count_m = thread.getMessageCount();
      for(var m=0;m<count_m;m++)
      {
        messages.push(thread.getMessages()[m]); 
      }
    }
  );
  
  return messages;
  
  // returns list of messages which are gmail-message class
  // https://developers.google.com/apps-script/reference/gmail/gmail-message
}

// --------------------------------------------
// Parse messages given a list of messages for Alta Estudiantes
// --------------------------------------------
function parseMessageData_altaEstudiantes(messages)
{
  // I usually receive many emails containing the same information
  // Example:
  // bla...bla...bla...bla...que Someone Surname (ssurname@XXX.edu) ...bla...bla..bla
  //
  // Using this function I parse the content of the email and return the information regarding 
  // someone's name and surname's, and email addres between brackets.
  
  var records=[];
  for(var m=0;m<messages.length;m++)
  {
    var text = messages[m].getPlainBody();
    //Logger.log(text)

    var matches = text.match(/que\s(.*)\s\*\((.*edu).*/);
    //var matches = text.match(/(.*)/);
    
    //if(!matches || matches.length < 1)
    //{
      //No matches; couldn't parse continue with the next message
    //  continue;
    //}
    var rec = {};
    rec.mail = matches[2];
    rec.name = matches[1];
    rec.date = messages[m].getDate();
    rec.id = "https://mail.google.com/mail/u/0/#inbox/" + messages[m].getId(); // Creates link to email
    rec.from_sender = messages[m].getFrom();

    //cleanup data    
    records.push(rec);
  }
  return records;
  // returns list of records containing some information
  // store for each email:
  // mail, name, date, id, sender
}

// --------------------------------------------
// Parse messages given a list of messages and regex
// --------------------------------------------
function parseMessageData_seguimiento(messages)
{
   // messages: list of gmail-message class objects
  var records=[];
  for(var m=0;m<messages.length;m++)
  {
    var rec = {};
    // ["sender_mail", "receiver_mail", "date", "email_id", "Subject"], 
    
    text_getFrom = messages[m].getFrom();
    rec.sender_email = text_getFrom;

    text_getTo = messages[m].getTo() + messages[m].getCc() + messages[m].getBcc();
    rec.receiver_email = text_getTo;

    rec.date = messages[m].getDate();
    rec.id = "https://mail.google.com/mail/u/0/#inbox/" + messages[m].getId(); // Creates link to email
    rec.subject = messages[m].getSubject();

    //cleanup data    
    records.push(rec);
  }
  return records;
  
  // returns list of records containing some information
  // store for each email:
  //"sender_mail", "receiver_mail", "date", "email_id", "Subject"
}

// --------------------------------------------
// Save information in Google sheets
// --------------------------------------------
function saveDataToSheet(records, header_table, spreadsheet_url, sheet_name)
{
  // records: array of objects with information to save
  // header_table: array of header column names
  // spreadsheet_url: URL for Google Spreadsheet to use
  // sheet_name: sheet name to store results

  var spreadsheet = SpreadsheetApp.openByUrl(spreadsheet_url);
  var sheet = spreadsheet.getSheetByName(sheet_name);
  
  // add header
  sheet.appendRow(header_table);
  
  // for each record save results
  for(var r=0;r<records.length;r++)
  {
    // Save values included in each object by creating an array of each element in the list records: Object.values(records[r])
    sheet.appendRow(Object.values(records[r])); 
  } 
  // Returns nothing.
  // Appends to given Google Sheet tab the given information  
}

// --------------------------------------------
// Function to process emails: Alta estudiantes
// --------------------------------------------
function processTransactionEmails_altaEstudiante(filter_str, header_table, spreadsheet_url, sheet_name)
{
  var messages = getRelevantMessages(filter_str);
  var records = parseMessageData_altaEstudiantes(messages);
  saveDataToSheet(records, header_table, spreadsheet_url, sheet_name);
  labelMessagesAsDone(messages);
}

// --------------------------------------------
// Function to process emails: Seguimiento
// --------------------------------------------
function processTransactionEmails_seguimiento(filter_str, header_table, spreadsheet_url, sheet_name)
{
  var messages = getRelevantMessages(filter_str);
  var records = parseMessageData_seguimiento(messages);
  saveDataToSheet(records, header_table, spreadsheet_url, sheet_name);
  labelMessages(messages, 'gap_parsed');
}

// --------------------------------------------
// Function to tag emails
// --------------------------------------------
function labelMessages(messages, label)
{
  // messages: list of messages
  // label: string to use as label
  var label_obj = GmailApp.getUserLabelByName(label);
  if(!label_obj)
  {
    label_obj = GmailApp.createLabel(label);
  }
  for(var m =0; m < messages.length; m++ )
  {
     label_obj.addToThread(messages[m].getThread() );  
  }
}

// --------------------------------------------
// Function to tag emails from a list of desired users
// --------------------------------------------
function tag_Messages_knonw_users(list_given, tags2tag)
{
  // list_given: wishlist of users that I want to follow
  // tags2tag: tags to tag each email (list or string)
  
  const users_list = list_given;
  //Logger.log(alumnos_list);

  // Get all new emails
  new_emails = getRelevantMessages('newer_than:24h AND -label:gap_parsed') // retrieve message not older than 1 day not previously parsed
  
  // parse each email to check if they belong to wishlist
  for(var m=0;m<new_emails.length;m++)
  {
    text_getFrom = new_emails[m].getFrom();
    Logger.log(text_getFrom);

    var matches_email_sender = text_getFrom.match(/.*\<(.*.edu)\>*/); // I check email comes from my instituion of interest
    Logger.log(matches_email_sender);
    
    // if emails comes from any of the user list provided
    if (!matches_email_sender) {
      continue;
    } else {
      if (users_list.includes(matches_email_sender[1])) { // I check sender is in my wishlist
        for (var t=0; t<tags2tag.length;t++) {
          labelMessages([ new_emails[m] ], tags2tag[t]); 
}}}}}
