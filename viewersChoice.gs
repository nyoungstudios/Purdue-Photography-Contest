// code written by Nathaniel Young
// @nyoungstudios on GitHub
// 04/28/2020

// Google document IDs
var votingSheetId = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
var submissionsSheetId2 = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';

// get sheet file
var votingSheet = SpreadsheetApp.openById(votingSheetId).getSheetByName('Sheet1');
var submissionsSheet = SpreadsheetApp.openById(submissionsSheetId2).getSheetByName('scriptTest');

// reads judging spreadsheet and creates dictionary
function createSubmissionsDict() {
  // sets initial variables
  var r = submissionsSheet.getDataRange();
  var values = r.getValues();
  
  // dictionary to store photo id mapped to email
  var submissionsId = {};
  
  //dictionary to store photo id mapped to name
  var submissionsName = {};
  
  // loops over all the rows and stores value in dictionaries
  for (var i = 1; i < values.length; i++) {
    var email = values[i][2];
    var name = values[i][1].toLowerCase();
    var photoId = values[i][7];
    submissionsId[photoId] = email;
    submissionsName[photoId] = name;
    
  }
  
//  Logger.log(submissionsId, submissionsName);
  
  return [submissionsId, submissionsName];
  
}

// function to calculate the scores for the viewer's choice
function calcScores() {
  // sets initial variables
  var r = votingSheet.getDataRange();
  var values = r.getValues();
  
  // scoring variables
  var scores = {};
  var scoringVector = [14, 9, 8, 7, 6, 5, 4, 3];
  var count = 0;
  
  // submission dictionaries
  var getSubmissionsDict = createSubmissionsDict(); 
  var submissionsId = getSubmissionsDict[0];
  var submissionsName = getSubmissionsDict[1];
  
  // sets for duplicate voters
  var nameSet = new Set();
  var emailSet = new Set();
  
  // loops over all rows except for title row
  for (var i = 1; i < values.length; i++) {
    var name = values[i][1].toLowerCase();
    var email = values[i][2];
    
    // checks to make sure this is their first vote ballot
    if (!nameSet.has(name) && !emailSet.has(email)) {
      nameSet.add(name);
      emailSet.add(email);
      
      // so you can't vote for the same photo
      var votedPhotosSet = new Set();
      
      // loops over all the votes in the columns
      for (var j = 3; j < values[i].length; j++) {
        var photoId = parsePhotoId(values[i][j]);
        //      Logger.log(photoId);
        
        // if the photo id is not blank, it is a valid photo id, the voter did not vote for their own photo, and didn't vote for a photo more than once
        if (photoId != '' && photoId in submissionsId && submissionsId[photoId] != email && submissionsName[photoId] != name && !votedPhotosSet.has(photoId)) {
          count++;
          votedPhotosSet.add(photoId);
          
          Logger.log(name, email, photoId);
                    
          if (photoId in scores) {
            scores[photoId] += scoringVector[j - 3];
          } else {
            scores[photoId] = scoringVector[j - 3];
          }
        }
        
      }
    }
  }
  
  // Create items array
  var sortedScores = Object.keys(scores).map(function(key) {
    return [key, scores[key]];
  });
  
  // Sort the array based on the second element
  sortedScores.sort(function(first, second) {
    return second[1] - first[1];
  });
  
  // print number of votes
  Logger.log('There are ' + count + ' votes.');
  
  // print sorted scores
  Logger.log(sortedScores);
  
}

// function to standardize the input string
function parsePhotoId(string) {
  string = string.toUpperCase();
  if (string.includes('-')) {
    return string.split(' - ')[0];
  } else {
    return string;
  }
}

