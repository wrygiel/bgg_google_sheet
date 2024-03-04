function extractGameIds(xmlString) {
  var document = XmlService.parse(xmlString);
  var root = document.getRootElement();
  var items = root.getChildren('item');
  var gameIds = [];
  
  for (var i = 0; i < items.length; i++) {
    var objectIdAttr = items[i].getAttribute('objectid');
    if (objectIdAttr) {
      gameIds.push(objectIdAttr.getValue());
    }
  }
  
  return gameIds;
}

function fetchGameIds() {
  var url = "https://boardgamegeek.com/xmlapi2/collection?username=wrygiel";
  var attempts = 0;
  var maxAttempts = 3;
  var success = false;
  var gameIds = [];
  
  while (attempts < maxAttempts && !success) {
    try {
      var response = UrlFetchApp.fetch(url);
      var xmlString = response.getContentText();
      gameIds = extractGameIds(xmlString);
      success = true;
    } catch (e) {
      console.error("Attempt " + (attempts + 1) + " failed: " + e.message);
      Utilities.sleep(5000);
      attempts++;
    }
  }
  
  if (!success) {
    throw new Error("Failed to fetch game IDs after " + maxAttempts + " attempts.");
  }
  
  return gameIds;
}

function dumpGameDataIntoSheet(xml) {
  var document = XmlService.parse(xml);
  var root = document.getRootElement();
  var items = root.getChildren('item');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Replaceable Game Dump");

  sheet.clearContents();
  // Header row
  var headers = [
    'ID', 'Name', 'Year Published', 'Thumbnail', 'Image', 'Playing Time', 'Min Play Time', 'Max Play Time', 'Min Age',
    'Language Dependence', ...Array.from({ length: 10 }, (_, i) => `${i + 1} Players`)
  ];
  sheet.appendRow(headers);
  
  items.forEach(item => {
    var rowData = [
      item.getAttribute('id').getValue(),
      item.getChild('name').getAttribute('value').getValue(),
      item.getChild('yearpublished').getAttribute('value').getValue(),
      item.getChild('thumbnail').getText(),
      item.getChild('image').getText(),
      item.getChild('playingtime').getAttribute('value').getValue(),
      item.getChild('minplaytime').getAttribute('value').getValue(),
      item.getChild('maxplaytime').getAttribute('value').getValue(),
      item.getChild('minage').getAttribute('value').getValue()
    ];

    var langDepPoll = item.getChildren('poll').find(poll => poll.getAttribute('name').getValue() === 'language_dependence');
    var langDepResults = langDepPoll.getChild('results').getChildren('result');
    var maxVotes = 0;
    var langDepValue = '';
    langDepResults.forEach(result => {
      var numVotes = parseInt(result.getAttribute('numvotes').getValue(), 10);
      if (numVotes > maxVotes) {
        maxVotes = numVotes;
        langDepValue = result.getAttribute('value').getValue();
      }
    });
    rowData.push(langDepValue);

    var numPlayersPoll = item.getChildren('poll').find(poll => poll.getAttribute('name').getValue() === 'suggested_numplayers');
    var playerCounts = {}; // To store the highest vote result for each player count
    var defaultValue = '-';
    numPlayersPoll.getChildren('results').forEach(resultsEl => {
      var numPlayers = resultsEl.getAttribute('numplayers').getValue();
      var mostVotedValue = "-";
      var mostVotedValueVotes = 0;
      resultsEl.getChildren('result').forEach(resultEl => {
        const votes = parseInt(resultEl.getAttribute('numvotes').getValue());
        if (votes > mostVotedValueVotes) {
          mostVotedValueVotes = votes;
          mostVotedValue = resultEl.getAttribute('value').getValue();
        }
      });
      if (`${numPlayers}` === `${parseInt(numPlayers)}`) {
        playerCounts[numPlayers] = mostVotedValue;
      } else if (`${numPlayers}` === `${parseInt(numPlayers)}+`) { // e.g. `2+`
        defaultValue = mostVotedValue;
      } else {
        throw new Error(`Unrecognized format: ${numPlayers}`);
      }
    });

    for (let i = 1; i <= 10; i++) {
      var playerCountValue = playerCounts[i] || defaultValue;
      rowData.push(playerCountValue);
    }

    sheet.appendRow(rowData);
  });
}

function dumpGamesIntoSheet(gameIds) {
  var url = "https://boardgamegeek.com/xmlapi2/thing?id=" + gameIds.join(",") + "&type=boardgame";
  var response = UrlFetchApp.fetch(url);
  var xmlString = response.getContentText();
  dumpGameDataIntoSheet(xmlString);
}

function run() {
  dumpGamesIntoSheet(fetchGameIds());
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('BGG')
      .addItem('Reload the sheet', 'run')
      .addToUi();
}
