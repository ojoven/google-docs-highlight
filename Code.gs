// Global var
var arraySplitters = ['.', ':' ,';', '?', '!'];

// On Open, creates a menu option and runs the show sidebar function
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Run', 'showSidebar')
      .addToUi();
}

// After install, open / run the script
function onInstall(e) {
  onOpen(e);
}

// Show sidebar
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Highlight');
  DocumentApp.getUi().showSidebar(ui);
}

function runDocumentTextHighlighter(threshold) {
  
  if (threshold === undefined) threshold = 5;
  var sentenceThresholdArray = getSentencesLengthBasedOnThreshold(threshold);
  Logger.log(sentenceThresholdArray);
  
  var body = DocumentApp.getActiveDocument().getBody();
  var paragraphs = body.getParagraphs();
  var stats = {};
  
  paragraphs.forEach(function(paragraph) {
    
    // Vars
    var pText = paragraph.editAsText();
    var pString = pText.getText();
    var sentencesArray = pString.split(/[.:;?!\r\n]+/);
    
    // Loop over the sentences
    sentencesArray.forEach(function(sentence) {
      
      var color = '#fff';
      
      // If not empty sentence
      if (sentence.trim()!="") {
        
        // Get the sentence's length
        var sentenceLength = sentence.trim().split(' ').length;
                
        if (sentenceLength <= sentenceThresholdArray[0]) {
          color = '#fff2cc'; // yellow
          stats = increaseNumStats(stats, 'yellow');
        } else if (sentenceLength <= sentenceThresholdArray[1]) {
          color = '#ffbff5'; // pink
          stats = increaseNumStats(stats, 'pink');
        } else if (sentenceLength <= sentenceThresholdArray[2]) {
          color = '#ff8f8f'; // red
          stats = increaseNumStats(stats, 'red');
        } else if (sentenceLength <= sentenceThresholdArray[3]) {
          color = '#aeffa8'; // green
          stats = increaseNumStats(stats, 'green');
        } else {
          color = '#8ff4ff'; // cyan
          stats = increaseNumStats(stats, 'cyan');
        }
      
        highlightSentence(paragraph, sentence, color);
      }
      
    });
    
  });
  
  // Return values to the HTML
  var response = {};
  response.stats = stats;
  response.sentenceLengths = sentenceThresholdArray;
  var responseString = JSON.stringify(response);
  return responseString;
  
}

function runDocumentTextHighlightCleaner() {
  var body = DocumentApp.getActiveDocument().getBody();
  var paragraphs = body.getParagraphs();
  
  paragraphs.forEach(function(paragraph) {
    
    // Vars
    var pText = paragraph.editAsText();
    var pString = pText.getText();
    var sentencesArray = pString.split('.');
    
    // Loop over the sentences
    sentencesArray.forEach(function(sentence) {
      
      var color = '#fff';
      
      // If not empty sentence
      if (sentence.trim()!="") {
        highlightSentence(paragraph, sentence, null);
      }
      
    });
    
  });
  
}

function increaseNumStats(stats, color) {
  if (typeof stats[color] == "undefined") {
   stats[color] = 1; 
  } else {
    stats[color]++; 
  }
  
  return stats;
}

function highlightSentence(paragraph, sentence, color) {

  var highlightStyle = {};
  highlightStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = color;
  var textLocation = {};
  var i;
  var paragraphText = paragraph.getText();

  textLocation = paragraph.findText(escapeRegExp(sentence));

  while (textLocation != null) {
        // Get the text object from the element
        var foundText = textLocation.getElement().asText();

        // Where in the Element is the found text?
        var indexStart = getHighlightIndexStart(textLocation, paragraphText);
        var indexEnd = getHighlightIndexEnd(textLocation, paragraphText);

        // Should we highlight the element found?
        if (shouldWeHighlight(paragraphText, indexEnd)) {

          // Change the background color to the belonging colour
          foundText.setAttributes(indexStart,indexEnd, highlightStyle);
        }

        // Find the next match
        textLocation = paragraph.findText(escapeRegExp(sentence), textLocation);
    }

  //Logger.log(paragraphText);
}

function shouldWeHighlight(paragraphText, indexEnd) {

  // We'll only highlight the found text if it ends in one of the splitting characters (.!?) or it's an end of line
  var charEnd = paragraphText[indexEnd];
  var charEndNext = paragraphText[indexEnd + 1];
  if (arraySplitters.indexOf(charEnd) !== -1 || charEndNext === undefined) {
    return true;
  }

  return false;
}

function getHighlightIndexStart(textLocation, paragraphText) {

  var indexStart = textLocation.getStartOffset();
  var charStart = paragraphText[indexStart];
  if (charStart == ' ') {
    indexStart++;
  }
  
  return indexStart;
}

function getHighlightIndexEnd(textLocation, paragraphText) {

  var indexEnd = textLocation.getEndOffsetInclusive();
  var charEndNext = paragraphText[indexEnd + 1];
  if (arraySplitters.indexOf(charEndNext) !== -1) {
   indexEnd++; 
  }
  
  return indexEnd;
}

function getSentencesLengthBasedOnThreshold(threshold) {
  
  var arrayThresholdLengths = [];
  arrayThresholdLengths[1] = [1,2,3,4];
  arrayThresholdLengths[2] = [1,2,4,6];
  arrayThresholdLengths[3] = [2,3,4,7];
  arrayThresholdLengths[4] = [2,3,5,8];
  arrayThresholdLengths[5] = [2,3,5,9];
  arrayThresholdLengths[6] = [2,4,6,10];
  arrayThresholdLengths[7] = [3,7,10,14];
  arrayThresholdLengths[8] = [5,9,14,18];
  arrayThresholdLengths[9] = [7,11,16,20];
  arrayThresholdLengths[10] = [10,14,18,24];
  arrayThresholdLengths[11] = [12,16,20,27];
  
  return arrayThresholdLengths[threshold];
  
}

function escapeRegExp(str) {
  return str.replace(/[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|]/g, "\\$&");
}