var globalEvent;

var sender = ""; //This will be set to the previous senders name
var senderEmail = ""; //This is the previous senders email

var currentSender = ""; //Current senders name
var currentEmail = ""; //Current senders email

var ss = SpreadsheetApp.openById("1CkXdLWZlPR5Ac0QSwB9kfDeq1zbOsQzmKn__QCqgIb4"); //Link to spreadsheet (the thing in the URL)
var sheet = ss.getSheets()[0]; //Returns the first sheet
var data = sheet.getDataRange().getValues(); 
//The previous line gets values of sheet as array (Be careful about this, the order is [y, x] instead of [x, y], and if you search for a cell that doesn't exist, it throws an error)

var currentSearchLink = ""; //Will become the search link to the conversations with the current email

var card = CardService.newCardBuilder(); //Creates a card - the side panel

function getContextualAddOn(event) { //Called when the button is pressed
  globalEvent = event;
  
  if(data[0][0] === ""){ //Checks if the spreadsheet file is empty, if yes, it resets the value
    resetInfo(); //Sets the sender and his email to the current sender and email
  }else{
    sender = data[0][0]; //Sets the previous sender to what was saved in spreadsheet file
    senderEmail = data[0][1]; //Same with email (here you can see the [y, x] since the sender and sender email are next to each other in the spreadsheet)
  }
  
  var currentNameAndEmail = getFrom(getCurrentMessage(event)); //Returns String in format: name <email>
  var dividedCurrent = currentNameAndEmail.split("<"); //Divides string at <, creating an array that looks like this: [name, <email>]
  currentSender = dividedCurrent[0]; //Sets currentSender = everything in front of <
  currentEmail = getCurrentMessage(event).getFrom().replace(/^.+<([^>]+)>$/, "$1"); //Honestly I don't know how this works but it does :D - Returns <email> without <>
  
  var currentSearchLink = "https://mail.google.com/mail/u/0/#search/" + currentEmail.replace("@", "%40"); 
  // Previous line creates the search link for the current email conversations (in the URL, there's a %40 instead of a @, that's why there's a replace)
  
  card.setHeader(CardService.newCardHeader().setTitle('Email Helper Thingy'));
  
  var action = CardService.newAction().setFunctionName('resetInfo'); //Button with onClickAction(action) does function resetInfo() when pressed - NOT USED
  
  var section = CardService.newCardSection(); //Creates a new sections within the card
  section.addWidget(CardService.newTextParagraph() //Adds text
     .setText("<i>Current Sender:</i>"));
  section.addWidget(CardService.newTextInput() //Adds textInput
    .setFieldName("Sender")
    .setValue(currentSender)); // The text inside the input box = currentSender
  section.addWidget(CardService.newTextInput()
    .setFieldName("SenderEmail")
    .setValue(currentEmail));
  section.addWidget(CardService.newTextButton() //Adds Button (Though it's like a link, not a button)
    .setText("Open Conversations")
    .setOpenLink(CardService.newOpenLink()
     .setUrl(currentSearchLink))); //Sets the link of the button to the currentSearchLink defined above
  section.addWidget(CardService.newTextParagraph()
    .setText("<i>Previous Sender:</i>"));
  section.addWidget(CardService.newTextInput()
    .setFieldName("Previous Sender")
    .setValue(sender)); // The text inside the input box = sender
  section.addWidget(CardService.newTextInput()
    .setFieldName("Previous Sender Email")
    .setValue(senderEmail));
  
  card.addSection(section);
  
  resetInfo(event); //Sets the previous sender to the current sender, though this doesn't display because the card isn't redrawn
  
  //NOTE: The card doesn't reset if you close and open it again, it only resets when you open a new email, thus the getContextualAddOn function is always only called once per email
  //NOTE 2: Bug - if you refresh the current page, the current and previous email are going to be the same... However, I don't think we'll use the previous email

  return card.build(); //Creates the interface
}

function getCurrentMessage(event) {
  var accessToken = event.messageMetadata.accessToken;
  var messageId = event.messageMetadata.messageId;
  GmailApp.setCurrentMessageAccessToken(accessToken);
  return GmailApp.getMessageById(messageId);
}

function getFrom(message) { //Returns the sender in the String: name <email>
  return message.getFrom();
}

function resetInfo(event){
  var tmpString = getFrom(getCurrentMessage(event)); //Returns String in format: name <email>
  var divided = tmpString.split("<"); //Divides string at <, creating an array that looks like this: [name, <email>]
  sender = divided[0]; //Sets sender = everything in front of <
  senderEmail = getCurrentMessage(event).getFrom().replace(/^.+<([^>]+)>$/, "$1"); //Honestly I don't know how this works but it does :D - Returns <email> without <>
  saveData(); //Saves data to the spreadsheet
}

function saveData() {
  sheet.clear(); //Removes everything from spreadsheet
  sheet.appendRow([sender, senderEmail]); //Adds the sender and senderEmail into first row of sheet
}
