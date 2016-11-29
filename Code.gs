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

function runDocumentTextHighlighter() {
  var body = DocumentApp.getActiveDocument().getBody();
  var paragraphs = body.getParagraphs();
  var stats = {};
  
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
      
        // Get the sentence's length
        var sentenceLength = sentence.trim().split(' ').length;
        if (sentenceLength <= 2) {
          color = '#fff2cc'; // yellow
          stats = increaseNumStats(stats, 'yellow');
        } else if (sentenceLength <= 4) {
          color = '#ffbff5'; // pink
          stats = increaseNumStats(stats, 'pink');
        } else if (sentenceLength <= 6) {
          color = '#ff8f8f'; // red
          stats = increaseNumStats(stats, 'red');
        } else if (sentenceLength <= 10) {
          color = '#aeffa8'; // green
          stats = increaseNumStats(stats, 'green');
        } else {
          color = '#8ff4ff'; // cyan
          stats = increaseNumStats(stats, 'cyan');
        }
      
        highlightText(paragraph, sentence, color);
      }
      
    });
    
  });
  
  var statsString = JSON.stringify(stats);
  return statsString;
  
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
        highlightText(paragraph, sentence, null);
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
  
  Logger.log(stats[color]);
  
  return stats;
}

function highlightText(paragraph, textToHighlight, color) {

  var highlightStyle = {};
  highlightStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = color;
  var textLocation = {};
  var i;

  textLocation = paragraph.findText(textToHighlight);
  if (textLocation != null && textLocation.getStartOffset() != -1) {
    textLocation.getElement().setAttributes(textLocation.getStartOffset(),textLocation.getEndOffsetInclusive(), highlightStyle);
  }
}