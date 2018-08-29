//==== Options

//If this is true, the script will only grab the first matching value it finds for each TrelloInput
//If this is false, it will fill more columns with all matching values on same-named cards
//Warning! Setting this to false will clear ALL CELLS to the right of each valid Input at every update.
var oneValuePerInput = true;

var trello_token = "paste_your_trello_token_here";
var trello_key = "paste_your_trello_api_key_here";
var trello_board_id = "paste_your_trello_board_id_here";


function onUserEdit(e) {
  doCompleteUpdate();
}

//A function required by Deploy as web app. It will be called when a webhook makes a POST request to our app.
function doPost(e) {
  try{
    var data = JSON.parse(e.postData.contents);
  } catch (err) {
    console.error( err );
    return HtmlService.createHtmlOutput("bad request");
  }
  
  if ( !filterAction( data.action ) ){
    return HtmlService.createHtmlOutput("ignoring this action");
  }
  
  doCompleteUpdate();
  
  return HtmlService.createHtmlOutput("post request received");
}


function doCompleteUpdate() {
  //Request ids of custom fields from Trello
  var customFields = getFieldsFromBoard( trello_board_id, trello_key, trello_token );
  //Get all cards
  var cards = getCards( trello_board_id, customFields, trello_key, trello_token );
  
  updateInputs( cards, customFields );
  var diff = getOutputValuesDifferentFromModel( cards, customFields );
  exportOutputs( diff, cards, customFields );
}


function filterAction( action ) {
  //We only need to watch for certain actions that might result in an update
  var acceptedActions = [ "updateCustomFieldItem", "addCustomField", "deleteCustomField",
                         "updateCustomField", "updateCard", "copyCard", 
                         "createCard", "convertToCardFromCheckItem", "deleteCard",
                         "moveCardToBoard", "moveCardFromBoard" ];
  if ( !action || acceptedActions.indexOf( action.type ) === -1 ){
    return false;
  }
  
  return true;
}


//Compares all values in TrelloOutput to a model and creates a  list of all differences.
function getOutputValuesDifferentFromModel( cards, customFields ) {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  
  var outputRange = doc.getRangeByName("TrelloOutput");
  var numRows = outputRange.getNumRows();
  
  var output = {};
  
  for (var i = 1; i <= numRows; i++) {
    var outputCell = outputRange.getCell(i, 1);
    var outputDefinition = outputCell.getValue().toLowerCase();
    var outputValue = outputCell.offset(0,1).getValue();
    
    var parsedDefinition = outputDefinition.split("@");
    
    //Definition is invalid
    if ( parsedDefinition.length < 2 ) continue;
    
    var cardName = parsedDefinition[0];
    var fieldName = parsedDefinition[1];
    
    //If we don't know a card or a field with such name, skip this cell
    if ( !cards[ cardName ] || !customFields.byName[ fieldName ] ) continue;
    
    //This script doesn't support List Custom Fields.
    if ( customFields.byName[ fieldName ].type === "list" ) continue;
    
    //Convert the value from spreadsheet to Trello field value for comparison with the model 
    var convertedValue = convertToTrelloValue( outputValue, customFields.byName[ fieldName ].type );
    if (convertedValue !== "") {
      convertedValue = convertedValue[Object.keys(convertedValue)[0]];
    } else {
      convertedValue = null;
    }
    
    //If this stays false, the value is the same as in the model
    var different = false;
    
    for (var j = 0; j < cards[ cardName ].length; j++) {
      var card = cards[ cardName ][j];
      
      if ( !( fieldName in card.customFields ) && convertedValue !== null && convertedValue !== "false"){
        different = true;
        break;
      } else if ( fieldName in card.customFields ) {
        if ( card.customFields[ fieldName ].value !== convertedValue ) {
          different = true;
          break;
        }
      }   
    }
    
    if (different) {
      output[ outputDefinition ] = outputValue;
    }
  }
  
  return output;
}


//Queries Trello API and updates every matched cell in TrelloInput
function updateInputs( cards, customFields ) {
  //Go over TrelloInput and update everything
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var inputRange = doc.getRangeByName("TrelloInput");
  var numRows = inputRange.getNumRows();
  
  for (var i = 1; i <= numRows; i++) {
    var inputCell = inputRange.getCell(i, 1);
    var inputDefinition = inputCell.getValue().split("@");
    
    //Definition is invalid
    if ( inputDefinition.length < 2 ) continue;
    
    if ( !oneValuePerInput ) {
      //Clear everything to the right of the cell
      var valuesRange = inputCell.offset( 0,1,1, sheet.getLastColumn() - inputCell.getColumn() );
      valuesRange.clearContent();
    }
    
    var cardName = inputDefinition[0].toLowerCase();
    var fieldName = inputDefinition[1].toLowerCase();
    
    //Anonymous function for getting an array of input values.
    //If oneValuePerInput is true, will always return an array of length 1
    var getValues = function() {
      var values = [];
      
      //If we don't know a card or a field with such name, skip this cell
      if ( !cards[ cardName ] || !customFields.byName[ fieldName ] ) return [""];
      
      var matchedCards = cards[ cardName ];
      
      for ( var j = 0; j < matchedCards.length; j++ ) {
        var card = matchedCards[j];
        
        if ( !card.customFields[ fieldName ] ) continue;
        
        //This script doesn't support List Custom Fields.
        if ( customFields.byName[ fieldName ].type === "list" ) continue;
        
        values.push( card.customFields[ fieldName ].value );
        if ( oneValuePerInput ) break;
      }
      
      if ( values.length === 0 ) values = [""];
      
      return values
    }
    
    var values = getValues();
    
    for ( var j = 0; j < values.length; j++ ) {
      inputCell.offset(0,1+j).setValue( values[j] );
    }
    
  }
  
  SpreadsheetApp.flush();
}


//Takes a dictionary of values indexed by 'Card Name@Field Name' definitions and exports them to Trello
function exportOutputs( outputs, cards, customFields ) {
  for (var outputDefinition in outputs) {
    
    var outputValue = outputs[ outputDefinition ];
    outputDefinition = outputDefinition.split("@");
    
    //Definition is invalid
    if ( outputDefinition.length < 2 ) continue;
    
    var cardName = outputDefinition[0].toLowerCase();
    var fieldName = outputDefinition[1].toLowerCase();
    
    //If we don't know a card or a field with such name, skip this value
    if ( !cards[ cardName ] || !customFields.byName[ fieldName ] ) continue;
    
    var matchedCards = cards[ cardName ];
    
    var fieldId = customFields.byName[ fieldName ].id;
    
    outputValue = convertToTrelloValue( outputValue, customFields.byName[ fieldName ].type );
    
    for ( var j = 0; j < matchedCards.length; j++ ){
      var cardId = matchedCards[j].id;
      setFieldOnCard( outputValue, cardId, fieldId, trello_key, trello_token );
    }
  }
}

//==== Trello API functions


//Queries Trello API and returns a dictionary of arrays of card data objects indexed by lowercased card names
//Arrays are used since multiple cards can share one name
function getCards( boardId, customFields, key, token ) {
  var url = "https://api.trello.com/1/boards/"+trello_board_id+"/cards/?fields=name&customFieldItems=true&" +
    "key=" + trello_key + "&token=" + trello_token;
  
  try {
    var response = UrlFetchApp.fetch( url );
  } catch (err) {
    console.error( err );
    return {};
  }
  
  var responseData = JSON.parse(response.getContentText());
  
  var cards = {};
  
  for (var i = 0; i < responseData.length; i++) {
    var card = responseData[i];
    
    var lowerName = card.name.toLowerCase();
    
    var customFieldItems = {};
    
    for ( var j = 0; j < card.customFieldItems.length; j++ ) {
      var fieldItem = card.customFieldItems[j];
      
      var fieldName = customFields.byId[ fieldItem.idCustomField ].name.toLowerCase();
      if ( !fieldName ) continue;
      
      var fieldValue = fieldItem.value;
      
      //This script doesn't support List Custom Fields.
      if ( customFields.byId[ fieldItem.idCustomField ].type === "list" ) {
        fieldValue = null; 
      }
      
      //If value is set, it will be an object with a single property.
      if (fieldValue !== null) {
        fieldValue = fieldValue[Object.keys(fieldValue)[0]];
      }
      
      customFieldItems[ fieldName ] = { id: fieldItem.id, name: fieldName, value: fieldValue };
    }
    
    var cardData = { id: card.id, name: card.name, customFields: customFieldItems };
    
    if ( cards[ lowerName ] ) {
      cards[ lowerName ].push( cardData );
      
    } else {
      cards[ lowerName ] = [ cardData ];
    }
  }
  
  return cards;
}



//Queries Trello API and returns two dictionary of same objects with custom fields data
//byId is a dictionary indexed by custom field ids
//byName is a dictionary indexed by lowercased custom field names
function getFieldsFromBoard( boardId, key, token ) {
  var url = "https://api.trello.com/1/boards/"+boardId+"/customFields?" + 
    "key=" + key + "&token=" + token;
  
  try {
    var response = UrlFetchApp.fetch( url );
  } catch (err) {
    console.error( err );
    return {};
  }
  
  var responseData = JSON.parse(response.getContentText());
  
  var customFields = { byName:{}, byId:{} };
  
  for (var i = 0; i < responseData.length; i++) {
    var field = responseData[i];
    var fieldData = { name: field.name, id: field.id, type: field.type };
    customFields.byName[ field.name.toLowerCase() ] = fieldData;
    customFields.byId[ field.id ] = fieldData;
  }
  
  return customFields;
}


//Sends an API request to change a value of a field on a Trello card
//value should be converted with convertToTrelloValue before calling this

function setFieldOnCard( value, cardId, fieldId, key, token ){
  
  var url = "https://api.trello.com/1/card/"+cardId+"/customField/"+fieldId+"/item";
  
  var payload = {
    value: value,
    key: key,
    token: token
  };
  
  var options = {
    contentType: "application/json",
    method: "put",
    payload: JSON.stringify(payload)
  };
  
  try {
    UrlFetchApp.fetch( url, options );
  } catch (err) {
    console.error( err )
  }
}


//Converts a value to an appropriate object acceptable by Trello API.
//type is a string taken from a custom field's definition returned by Trello.
//can be 'string', 'number', 'checkbox', 'date', 'list'

function convertToTrelloValue( value, type ) {
  
  //Clear the field if the value is empty
  if ( value === "" || value === null || value === undefined ) {
    return ""; 
  }
  
  switch (type) {
    case "text":
      return { text: value.toString() };
      
    case "number":
      var n = parseFloat(value);
      if ( isNaN(n) ) {
        return "";
      } else {
        return { number: n.toString() };
      }
      
    case "checkbox":
      return { checked: (!!value).toString() };
      
    case "date":
      var d = new Date(value);
      if ( isNaN( d.getTime() ) ) {
        return "";
      } else {
        return { date: d.toISOString() };
      }
      
    case "list":
      //This script doesn't support List Custom Fields.
      return "";
      
    default:
      //This shouldn't happen. We can't assume the type, so we clear the field instead.
      return "";
  }
}


//Register a webhook
function register() {
  var url = "https://api.trello.com/1/tokens/"+trello_token+"/webhooks/?key="+trello_key;
  
  var payload = {
    description: "Google Spreadsheet webhook",
    callbackURL: ScriptApp.getService().getUrl(),
    idModel: trello_board_id
  };
  
  var options = {
    method: "post",
    payload: payload
  };
  
  UrlFetchApp.fetch( url, options );
}


//Unregister all webhooks with callbackURL of this app
function unregister() {
  var url = "https://api.trello.com/1/tokens/"+trello_token+"/webhooks?key="+trello_key;
  
  var response = UrlFetchApp.fetch( url );
  
  var responseData = JSON.parse(response.getContentText());
  
  for ( var i = 0; i < responseData.length; i++ ) {
    var webhook = responseData[i];
    
    if ( webhook.callbackURL !== ScriptApp.getService().getUrl() ) continue;
    
    var url = "https://api.trello.com/1/tokens/"+trello_token+"/webhooks/"+webhook.id+"?key="+trello_key;
    
    var response = UrlFetchApp.fetch( url, { method: "delete" } );
  }
}
