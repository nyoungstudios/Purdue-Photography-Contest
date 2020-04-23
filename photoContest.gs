// code written by Nathaniel Young
// @nyoungstudios on GitHub
// 04/22/2020

// Google document IDs
var sheetId = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
var docId1 = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
var docId2 = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';

// get sheet file
var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('scriptTestSend');

// function to send one email
function emailEntry(email, name, photoTitle, originalityScore1, aestheticsScore1, technicalityScore1, totalScore1, originalityScore2, aestheticsScore2, technicalityScore2, totalScore2, finalScore, judge1Comments, judge2Comments) {
  // gets email alias
  var aliases = GmailApp.getAliases();
  //Logger.log(aliases[0]);
  
  // test 
//  var me = Session.getActiveUser().getEmail();
  
  var subject = 'Purdue Photography Contest Scores for your photo: ' + photoTitle;
  
  var body1 = 'Hi ' + name + ','
  
  var body2 = 'Thank you for entering the Purdue Photography Contest! We had a lot of entries this year, and the judges were \
really impressed with the quality of everyone\'s photos. Here are your scores for your photo titled: ' + photoTitle + '.';

  var body3 = scoreTable('first', originalityScore1, aestheticsScore1, technicalityScore1, totalScore1);
  
  var body4 = 'Comments: ' + judge1Comments;
  
  var body5 = scoreTable('second', originalityScore2, aestheticsScore2, technicalityScore2, totalScore2);
  
  var body6 = 'Comments: ' + judge2Comments;
  
  var body7 = 'Average Total Score (out of 100): ' + finalScore;
  
  var body8 = 'Also, the photo gallery of all the submissions is open for viewing in addition to the viewer\'s choice on our website. \
Check out our email newsletter we sent out today announcing the winners of our photo contest and how to vote for the viewer\'s choice. \
Or visit this link for the same details: [URL HERE]';
  
  var body9 = 'Best regards,\nThe Purdue Photography Club';
  
  var body = buildBody(body1, body2, body3, body4, body5, body6, body7, body8, body9);
  
  GmailApp.sendEmail(email, subject, body, {'from': aliases[0]});
}

// function to create score table
function scoreTable(number, originalityScore, aestheticsScore, technicalityScore, totalScore) {
  var text1 = 'Scoring from our ' + number + ' judge.\n';
  var text2 = 'Originality (out of 25): ' + originalityScore + '\nAesthetics (out of 50): ' + aestheticsScore + '\nTechnicality (out of 25): ' + 
   technicalityScore + '\nTotal (out of 100): ' + totalScore;
  return text1 + text2;
}

// function to build the body of the email by adding new lines between paragraphs
function buildBody() {
  var body = '';
  var spaceText = '\n\n';
  
  for (var i = 0; i < arguments.length; i++) {
    if (i < arguments.length - 1) {
      body += arguments[i] + spaceText;
    } else {
      body += arguments[i];
    }
  }
  
  return body;
}

// function to read the scoring document
function readDocs(judgeNumber) {
  var doc;
  if (judgeNumber == 1) {
    doc = DocumentApp.openById(docId1);
  } else {
    doc = DocumentApp.openById(docId2);
  }
  
  
  // Define the search parameters.
  var searchTable = DocumentApp.ElementType.TABLE;
  var searchHeading = DocumentApp.ParagraphHeading.HEADING1;
  var searchResult = null;
  
  
  var submissionNumber = 0;
  
  var commentFlag = false;
  
  var photoId;
  var originalityScore;
  var aestheticsScore;
  var technicalityScore;
  var totalScore;
  var comment;
  
  for (var i = 50; i < doc.getNumChildren(); i++) {
    // breaks for testing
//    if (submissionNumber > 50) {
//      break;
//    }
    
    // gets item
    var item = doc.getChild(i);
    
    if (item.getType() == DocumentApp.ElementType.PARAGRAPH) {
      // tests if the item is a paragraph
      // then parses the paragraph
      if (item.getText().startsWith('--------------------------------')) {
        commentFlag = false;
//        Logger.log(photoId);
//        Logger.log(originalityScore, aestheticsScore, technicalityScore, totalScore);
//        Logger.log(comment);
        
        // creates value object
        var values = [[originalityScore, aestheticsScore, technicalityScore, totalScore, comment]];
        
        // input data in spreadsheet
        inputData(judgeNumber, submissionNumber, photoId, values);
        
        submissionNumber++;
                  
        var photoId = '';
        var originalityScore = 0;
        var aestheticsScore = 0;
        var technicalityScore = 0;
        var totalScore = 0;
        var comment = '';
      } else if (item.getText().startsWith('Photo ID: ')) {
        photoId = item.getText().split('Photo ID: ')[1].split(' - ')[0];
      } else if (item.getText().startsWith('Comments:')) {
        var temp = item.getText().split('Comments: ');
        if (temp.length > 1) {
          comment = temp[1];
        }
      }
    } else if (item.getType() == DocumentApp.ElementType.TABLE) {
      // tests if the item is a table
      // then parses the table
      originalityScore = parseInt(item.getCell(1, 1).getText());
      aestheticsScore = parseInt(item.getCell(2, 1).getText());
      technicalityScore = parseInt(item.getCell(3, 1).getText());
      totalScore = originalityScore + aestheticsScore + technicalityScore;
    }

    
  }
  
  
  Logger.log('Number of submissions: ' + submissionNumber);
  
  
}

// function to input scoring data into sheets
function inputData(judgeNumber, submissionNumber, photoId, values) {  
  
  // sets initial variables
  var row = submissionNumber + 2;
  var r;
  if (judgeNumber == 1) {
    r = sheet.getRange('J' + row + ':N' + row);
  } else {
    r = sheet.getRange('O' + row + ':S' + row);
  }
    
  // get photo id cell
  var c = sheet.getRange(row, 8);
  
  // verify it is adding the value to the right person
  if (c.getValue() == photoId) {
    r.setValues(values)
  } else {
    Logger.log('Failed on: ' + row);
  }
    
  
}

// function to call readDocs and update the values in the sheets file from both judges
function readAndUpdate() {
  readDocs(1);
  readDocs(2);
}

// function to read the sheets file and call the email function
function readSheetsAndEmail() {
  
  // sets initial variables
  var r = sheet.getDataRange();
  var values = r.getValues();
  
  // loops over all rows except for title row
  for (var i = 1; i < values.length; i++) {
    // gets values
    var email = values[i][2];
    var name = values[i][1];
    var photoTitle = values[i][5];
    var originalityScore1 = values[i][9];
    var aestheticsScore1 = values[i][10];
    var technicalityScore1 = values[i][11];
    var totalScore1 = values[i][12];
    var originalityScore2 = values[i][14];
    var aestheticsScore2 = values[i][15];
    var technicalityScore2 = values[i][16];
    var totalScore2 = values[i][17];
    var finalScore = values[i][19];
    var judge1Comments = values[i][13];
    var judge2Comments = values[i][18];
    
    // prints values to log
    Logger.log('email: ' + email);
    Logger.log('name: ' + name);
    Logger.log('photo title: ' + photoTitle);
    Logger.log(originalityScore1, aestheticsScore1, technicalityScore1, totalScore1);
    Logger.log(originalityScore2, aestheticsScore2, technicalityScore2, totalScore2);
    Logger.log(finalScore);
    Logger.log(judge1Comments);
    Logger.log(judge2Comments);
    Logger.log('--------------------------------');
    
    // call send email function
    emailEntry(email, name, photoTitle, originalityScore1, aestheticsScore1, technicalityScore1, totalScore1, originalityScore2, aestheticsScore2, technicalityScore2, totalScore2, finalScore, judge1Comments, judge2Comments);

  }
  
  
}


function test() {
  emailEntry('', 'Nathaniel Young', 'Photo Title', 25, 50, 25, 100, 22, 40, 16, 78, 89, 'Test comment numbet 1 here.', 'Other more different test comment.')
}



