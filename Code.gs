/**
* The onOpen function runs automatically when the Google Docs document is
* opened. Use it to add custom menus to Google Docs that allow the user to run
* custom scripts. For more information, please consult the following two
* resources.
*
* Extending Google Docs developer guide:
*     https://developers.google.com/apps-script/guides/docs
*
* Document service reference documentation:
*     https://developers.google.com/apps-script/reference/document/
*/
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
  .addItem('Open Healthcare Sidebar', 'showSidebar')
  .addToUi();
}

/**
* Shows a custom HTML user interface in a sidebar in the Google Docs editor.
*/
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('My_sidebar').setTitle('Healthcare sidebar').setWidth(300);
  DocumentApp.getUi().showSidebar(html);
  var default_types = ['observation', 'action', 'title', 'information', 'author', 'diagnosis'];
  var default_colors = ['#EDE96F', '#FF0000', '#80CF83', '#64A5FA', '#C781EB', '#E8B890'];
  var all_types = new Array();
  var lastID = PropertiesService.getDocumentProperties().getProperty('ID');
  if(lastID == null) {
    PropertiesService.getDocumentProperties().setProperty('ID', 0);
  }
  for (var i = 0; i < default_types.length; i++) {
    all_types[i] = new Object();
    all_types[i].name = default_types[i];
    all_types[i].color = default_colors[i];
    all_types[i].attributes = new Array();
    
    all_types[i].attributes[0] = new Object();
    all_types[i].attributes[0].name = 'id';
    all_types[i].attributes[1] = new Object();
    all_types[i].attributes[1].name = 'attribut_'+all_types[i].name;
  }
  clearCache();
  initCache(all_types);
}

/**
* Generates a Google Spreadsheet from the document
*/
function generateSpreadsheet() {
  var fileName = DocumentApp.getActiveDocument().getName()+"_data";
  var spreadsheet = SpreadsheetApp.create(fileName); 
  var documentProperties = PropertiesService.getDocumentProperties();
  var annotations = JSON.parse(documentProperties.getProperty('ANNOTATIONS'));
  var annotations_type = JSON.parse(documentProperties.getProperty('ANNOTATIONS_TYPE'));  
  var links = JSON.parse(documentProperties.getProperty('LINKS'));  
  var types =  Object.keys(annotations_type);
  
  /* First type */  
  spreadsheet.getActiveSheet().setName(types[0]);
  for (var i = 0; i < annotations_type[types[0]].attributes.length; i++) {
    var attributeName = annotations_type[types[0]].attributes[i].name;
    spreadsheet.getSheetByName(types[0]).getRange(1,i+1).setValue(attributeName);
    for (var j = 0; j < annotations[types[0]].length; j++) {     
      spreadsheet.getSheetByName(types[0]).getRange(j+2,1+i).setValue(annotations[types[0]][j][attributeName]);
    }
  }
  /* Saving of the links in the last column */
  var indexLastColumn = annotations_type[types[0]].attributes.length;
  spreadsheet.getSheetByName(types[0]).getRange(1,indexLastColumn+1).setValue("links");
  for (var j = 0; j < annotations[types[0]].length; j++) {    
    var currentID = annotations[types[0]][j]["id"];
    if(links[currentID] != undefined) {
      var linksToString = "";
      for (var k = 0; k < links[currentID].length; k++) {
        if(linksToString == "") {
          linksToString = links[currentID][k]["target"];
        }
        else {
          linksToString = linksToString +"," + links[currentID][k]["target"];
        }
      }
      spreadsheet.getSheetByName(types[0]).getRange(j+2,1+indexLastColumn).setValue(linksToString);
    }
  }  
  
  /* Other types */
  for (var i = 1; i < types.length; i++) {
    spreadsheet.insertSheet(types[i]);
    for (var j = 0; j < annotations_type[types[i]].attributes.length; j++) {  
      var attributeName = annotations_type[types[i]].attributes[j].name;
      spreadsheet.getSheetByName(types[i]).getRange(1,j+1).setValue(attributeName);
      for (var k = 0; k < annotations[types[i]].length; k++) {       
        spreadsheet.getSheetByName(types[i]).getRange(k+2,1+j).setValue(annotations[types[i]][k][attributeName]);
      }
    }
    /* Saving of the links in the last column */
    var indexLastColumn = annotations_type[types[i]].attributes.length;
    spreadsheet.getSheetByName(types[i]).getRange(1,1+indexLastColumn).setValue("links");
    for (var j = 0; j < annotations[types[i]].length; j++) {    
      var currentID = annotations[types[i]][j]["id"];
      if(links[currentID] != undefined) {
        var linksToString = "";
        for (var k = 0; k < links[currentID].length; k++) {
          if(linksToString == "") {
            linksToString = links[currentID][k]["target"];
          }
          else {
            linksToString = linksToString +"," + links[currentID][k]["target"];
          }
        }
        spreadsheet.getSheetByName(types[i]).getRange(j+2,1+indexLastColumn).setValue(linksToString);
      }
    }
  }
}

/* Adds a type to the selected text */
function save(type, color) {   
  var selectedElements = getSelectedText();
  var annotation = addToCache(type, selectedElements.text, selectedElements.elements);
  updateView(selectedElements.elements, color);
  insertComment(DocumentApp.getActiveDocument().getId(), selectedElements.text, annotation);
}

/* Adds the selected text to the cache - returns the new annotation as a String from the JSON object */
function addToCache(type, text, selectedElements) {
  var annotations = JSON.parse(CacheService.getPrivateCache().get("annotations"));
  var annotations_type = JSON.parse(PropertiesService.getDocumentProperties().getProperty('ANNOTATIONS_TYPE'));
  var lastID = PropertiesService.getDocumentProperties().getProperty('ID');
  var annotation = {};
  for(i = 0; i< annotations_type[type].attributes.length; i++) {
    var attrName = annotations_type[type].attributes[i].name;
    var nbChar = annotations_type[type].attributes[i].char; 
    var pattern = annotations_type[type].attributes[i].pattern; 
    var overridable = annotations_type[type].attributes[i].overridable; 
    var manual = annotations_type[type].attributes[i].manual; 
    if(pattern != null) {
      text = findPattern(selectedElements, pattern);
    }
    if(overridable != undefined && overridable == "true") {    
      var ui = DocumentApp.getUi();
      var response = ui.prompt('Override '+attrName, 'Override the value "'+text+'"' + '?', ui.ButtonSet.OK_CANCEL);
      if(response.getSelectedButton() == ui.Button.OK)
        text = response.getResponseText();
    }
    if(manual != undefined && manual == "true") {    
      var ui = DocumentApp.getUi();
      var response = ui.prompt('Manual attribute', 'Value for ' + attrName + '?', ui.ButtonSet.OK);
      text = response.getResponseText();
    }
    if(nbChar != null)
      annotation[attrName] = ""+text.substring(0, nbChar); // to avoid the [] 
    else {
      if(attrName == 'id') {
        annotation[attrName] = ""+lastID++; // to avoid the [] 
        PropertiesService.getDocumentProperties().setProperty('ID', lastID);
      }
      else
        annotation[attrName] = ""+text; // to avoid the []    
    }
  }
  annotations[type].push(annotation);  
  CacheService.getPrivateCache().put('annotations', JSON.stringify(annotations), 3600);
  return JSON.stringify(annotation);
}

/* Finds and returns the pattern within the selected elements */
function findPattern(selectedElements, pattern) { 
  for (var i = 0; i < selectedElements.length; ++i) {
    var element = selectedElements[i];
    var matchingElem; var value = '';
    if (element.isPartial()) {
      DocumentApp.getUi().alert('Warning: the search for the pattern will apply to the whole paragraph.'); 
    }
    // else {
    matchingElem = element.getElement().asText().findText(pattern);
    if(matchingElem != null) {
      var startIndex = matchingElem.getStartOffset();
      var endIndex = matchingElem.getEndOffsetInclusive();
      value = matchingElem.getElement().asText().getText().substring(startIndex, endIndex+1);
    }
    //   }
  }
  // NB: returns the last matching element
  return value;
}

/* Highlights the selected text */
function updateView(selectedElements, backgroundColor) { 
  for (var i = 0; i < selectedElements.length; ++i) {
    var element = selectedElements[i];
    if (element.isPartial()) {
      var startIndex = element.getStartOffset();
      var endIndex = element.getEndOffsetInclusive();
      element.getElement().asText().setBackgroundColor(startIndex, endIndex, backgroundColor);
    }
    else {
      element.getElement().asText().setBackgroundColor(backgroundColor);
    }
  }
}

/* Initialise the cache and the document properties */
function initCache(xml_content) {
  var annotations = {}; 
  var annotations_type = {}; 
  var links = {}; 
  for (var i = 0; i < xml_content.length; i++) {
    annotations[xml_content[i].name] = [];
    annotations_type[xml_content[i].name] = {"color" : xml_content[i].color, "attributes" : new Object(xml_content[i].attributes)};
  }  
  CacheService.getPrivateCache().put('annotations', JSON.stringify(annotations), 3600);   
  CacheService.getPrivateCache().put('links', JSON.stringify(links), 3600);   
  PropertiesService.getDocumentProperties().setProperty('ANNOTATIONS_TYPE', JSON.stringify(annotations_type));
}

/* Clears the annotations stored into the cache */
function clearAnnotationsInCache() {
  if(CacheService.getPrivateCache().get('annotations') != null) {
    var annotations = JSON.parse(CacheService.getPrivateCache().get('annotations'));
    var types =  Object.keys(annotations);
    var emptyAnnotations = {}; 
    for (var i = 0; i < types.length; i++) {
      emptyAnnotations[types[i]] = [];
    }  
    CacheService.getPrivateCache().put('annotations', JSON.stringify(emptyAnnotations), 3600);
  }
}

/* Clears the links stored into the cache */
function clearLinksInCache() {
  CacheService.getPrivateCache().put('links', JSON.stringify({}), 3600);
}

/* Clears the cache */
function clearCache() {
  clearAnnotationsInCache();
  clearLinksInCache();
}  

/* Saves the annotations in the document properties */
function saveAnnotationsInDoc() {
  var annotationsInCache = JSON.parse(CacheService.getPrivateCache().get('annotations'));
  var documentProperties = PropertiesService.getDocumentProperties();
  var annotationsInDoc = documentProperties.getProperty('ANNOTATIONS');
  var annotations_type = JSON.parse(documentProperties.getProperty('ANNOTATIONS_TYPE'));
  var manualAnnotations = retrieveAnchoredComments();
  
  if(annotationsInDoc == null) {
    documentProperties.setProperty('ANNOTATIONS', JSON.stringify(annotationsInCache));
    clearAnnotationsInCache();
    saveAnnotationsInDoc(); // for the manual annotations
  }
  else {
    annotationsInDoc = JSON.parse(documentProperties.getProperty('ANNOTATIONS'));
    var newAnnotations = annotationsInDoc;
    var keys = Object.keys(annotationsInDoc);
    for(var k in keys) {
      var type = keys[k];
      /* Save annotations from the cache into the newAnnotations object */
      if(annotationsInCache[type].length > 0) {
        for(var i = 0; i < annotationsInCache[type].length; i++) {
          var annotation = {};
          for(j = 0; j< annotations_type[type].attributes.length; j++) {
            var attrName = annotations_type[type].attributes[j].name;
            annotation[attrName] = ""+annotationsInCache[type][i][attrName];  
          }
          newAnnotations[type].push(annotation);
        }
      }
      /* Save manual annotations from the comments into the newAnnotations object */
      if(manualAnnotations[type].length > 0) {
        for(var i = 0; i < manualAnnotations[type].length; i++) {
          var annotation = {};
          for(j = 0; j< annotations_type[type].attributes.length; j++) {
            var attrName = annotations_type[type].attributes[j].name;
            annotation[attrName] = ""+manualAnnotations[type][i][attrName];  
          }
          newAnnotations[type].push(annotation);
        }
      }
    }
    documentProperties.setProperty('ANNOTATIONS', JSON.stringify(newAnnotations));
    clearAnnotationsInCache();
  }
}

/**
* Insert a new document-level comment.
*
* @param {String} fileId ID of the file to insert comment for.
* @param {String} content Text content of the comment.
*/
function insertComment(fileId, selection, type) {
  var comment = Drive.newComment();
  var context = Drive.newCommentContext();
  comment.content = type; 
  context.value = selection;
  context.type = 'text/html';
  comment.context = context;
  Drive.Comments.insert(comment, fileId) ;  
}

/* Displays the annotations stored in the document properties */
function getAnnotationsInDoc() {
  var documentProperties = PropertiesService.getDocumentProperties();
  DocumentApp.getUi().alert('Annotations in the Document Properties: ' + documentProperties.getProperty('ANNOTATIONS'));
  //DocumentApp.getUi().alert('Annotations_type in the Document Properties: ' + documentProperties.getProperty('ANNOTATIONS_TYPE'));  
  //DocumentApp.getUi().alert('Links in the Document Properties: ' + documentProperties.getProperty('LINKS'));  
}

/* Remove the annotations stored in the document properties */
function clearAnnotationsInDoc() {
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.deleteProperty('ANNOTATIONS');
  documentProperties.deleteProperty('ANNOTATIONS_TYPE'); 
  documentProperties.deleteProperty('ID'); 
  documentProperties.deleteProperty('LINKS'); 
  // A reload of the page is needed after this function, otherwise the types are empty
}

/* Parses the XML file and customise the application with the new types */
function loadXML(e) {
  var url = e.parameter.url;
  var xml = UrlFetchApp.fetch(url).getContentText();
  var document = XmlService.parse(xml); 
  var entries = document.getRootElement().getChildren();
  var all_types = new Array();
  
  for (var i = 0; i < entries.length; i++) {
    all_types[i] = new Object();
    all_types[i].name = entries[i].getChildText('name');
    all_types[i].color = entries[i].getChildText('color');
    all_types[i].attributes = new Array();
    var children =  entries[i].getChildren('attribute');
    for (var j = 0; j < children.length; j++) {
      all_types[i].attributes[j] = new Object();
      all_types[i].attributes[j].name = children[j].getText();
      var char = children[j].getAttribute('char');
      if(char != null)
        all_types[i].attributes[j].char = char.getValue();
      var pattern = children[j].getAttribute('pattern');
      if(pattern != null)
        all_types[i].attributes[j].pattern = pattern.getValue();
      var overridable = children[j].getAttribute('overridable');
      if(overridable != null && overridable.getValue() == 'true') {
        all_types[i].attributes[j].overridable = overridable.getValue();        
      }
      var manual = children[j].getAttribute('manual');
      if(manual != null && manual.getValue() == 'true') {
        all_types[i].attributes[j].manual = manual.getValue();      
      }
    }    
  }
  
  updateButtons(all_types);
  initCache(all_types);
  DocumentApp.getUi().alert('Sidebar customised');
}

/**
* Gets the text the user has selected. If there is no selection,
* this function displays an error message.
*
* @return {Array.<string>} The selected text.
*/
function getSelectedText() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  
  if (!selection) {
    DocumentApp.getUi().alert('Cannot find a selection in the document.');
    return;
  }
  
  var struct_elements = new Array();  
  struct_elements.elements = new Array();  
  var elements = selection.getSelectedElements();
  for (var i = 0; i < elements.length; i++) {
    if (elements[i].isPartial()) {
      var element = elements[i].getElement().asText();
      var startIndex = elements[i].getStartOffset();
      var endIndex = elements[i].getEndOffsetInclusive();
      struct_elements.elements[i] = elements[i];
      if(struct_elements.text == undefined)
        struct_elements.text = element.getText().substring(startIndex, endIndex + 1); 
      else
        struct_elements.text = struct_elements.text + element.getText().substring(startIndex, endIndex + 1);
    } else {
      var element = elements[i].getElement();
      // Only translate elements that can be edited as text; skip images and
      // other non-text elements.
      if (element.editAsText) {
        var elementText = element.asText().getText();
        // This check is necessary to exclude images, which return a blank text element.
        if (elementText != '') {    
          struct_elements.elements[i] = elements[i];
          if(struct_elements.text == undefined) 
            struct_elements.text = '' ; // Without this the word 'undefined' will be written in struct_elements.text
          struct_elements.text = struct_elements.text + elementText;         
        }
      }
    }
  }
  if (struct_elements.text.length == 0) {
    throw 'Please select some text.';
  }
  return struct_elements;  
}

/* Popup for customising the sidebar */
function customise() {
  var app = UiApp.createApplication();  
  
  var form = app.createFormPanel().setId('form').setEncoding('multipart/form-data');
  app.add(form);
  
  //a grid to hold the widgets to guild user
  var formContent = app.createGrid().resize(2,3);
  form.add(formContent);
  
  formContent.setText(0, 0, 'Use an existing file: ')
  formContent.setWidget(0, 1, app.createTextBox().setName('url'));
  
  var clickHandler = app.createServerHandler('loadXML');
  clickHandler.addCallbackElement(form);
  formContent.setWidget(0, 2, app.createButton('Customise').addClickHandler(clickHandler));  
  
  DocumentApp.getUi().showModalDialog(app, 'XML file');
}

/* Displays the customised sidebar */
function updateButtons(xml_content) {
  var newSidebar = '<link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons.css">';
  newSidebar += '<link rel="stylesheet" href="https://raw.githubusercontent.com/UOY-Enterprise/gdoc-annotations-addon/master/style_healthcare.css">';
  
  newSidebar +=('<div class="sidebar main">  <div class="step gettingStarted">    <b>Getting Started</b><br> </div>');
  newSidebar += '<div class="selectText">      Select the text you want to annotate    </div>';
  newSidebar += '<div class="step selectColor">   <b>Type</b>   <br>Choose a type for your selection';
  newSidebar += '<div class="types_buttons"> <table>';
  for (var i = 0; i < xml_content.length; i++) {
    newSidebar += '<tr><button onclick="save(this)" class="my_button" name="'+xml_content[i].name+'" value="'+xml_content[i].color+'">' +xml_content[i].name+'</button></tr>';
  }  
  
  newSidebar +='</table></div>';
  newSidebar += '<div class="links_buttons"><table>';
  newSidebar += '<tr><button onclick="google.script.run.displayLinkCreation()" class="my_button" name="link_creator">Link creator</button></tr>';
  newSidebar += '<tr><button onclick="google.script.run.showLinks()" class="my_button" name="link_creator">Show links</button></tr>';  
  newSidebar += '</table></div>';  
  
  newSidebar += ' <table> <div class="other_buttons"> <tr><button onclick="google.script.run.getAnnotationsInDoc()">What\'s in the doc</button>';
  newSidebar += '<button onclick="google.script.run.saveAnnotationsInDoc()" class="action">Save</button></tr>';
  
  newSidebar += '<tr><button onclick="google.script.run.clearAnnotationsInDoc()">Clear doc prop</button>';
  newSidebar += '<button onclick="google.script.run.generateSpreadsheet()" class="action">Generate Spreadsheet</button></tr>';
  
  newSidebar += '<tr><button onclick="google.script.run.clearCache()">Clear cache</button></tr>';
  // newSidebar += '<button onclick="google.script.run.customise()">customise</button>';
  
  newSidebar += '</div></table>';
  
  newSidebar +='</div> </div>';
  newSidebar +='<script> function save(button) {google.script.run.save(button.name, button.value);}</script>';
  var htmlOutput = HtmlService.createHtmlOutput(newSidebar);
  htmlOutput.setWidth(250).setHeight(300);
  DocumentApp.getUi().showSidebar(htmlOutput);
}


/* Window for creating links between elements */
function displayLinkCreation() {
  var app = UiApp.createApplication().setWidth(600).setHeight(450);  
  var panelLeft = app.createVerticalPanel();
  var panelRight = app.createVerticalPanel(); 
  var horPanel = app.createHorizontalPanel();
  var scrollPanel = app.createScrollPanel().setPixelSize(595, 390);
  var mainPanel = app.createVerticalPanel();    
  var formPanel = app.createFormPanel(); 
  
  panelLeft.add(app.createLabel("Source"));
  panelRight.add(app.createLabel("Target"));
  var allAnnotations = JSON.parse(PropertiesService.getDocumentProperties().getProperty('ANNOTATIONS'));
  var types =  Object.keys(allAnnotations);
  for (var i = 0; i < types.length; i++) {
    annotations = allAnnotations[types[i]];
    if(annotations.length != 0) {
      var attributes = Object.keys(annotations[0]);
      if(attributes.length > 0) {     
        for (var j = 0; j < annotations.length; j++) {
          var annotation = annotations[j];      
          // the default value for the text to print will be the second attribute
          panelLeft.add(  app.createRadioButton('radio_src',    annotation[attributes[1]]).setFormValue(annotation.id) );
          panelRight.add( app.createRadioButton('radio_target', annotation[attributes[1]]).setFormValue(annotation.id) );
        }
      }
    }
  }
  
  var createButton = app.createSubmitButton('Create link');
  
  var doneButton = app.createButton('Done');
  var clickHandler = app.createServerHandler('persistLinks');
  clickHandler.addCallbackElement(mainPanel);
  doneButton.addClickHandler(clickHandler);
  
  horPanel.add(panelLeft);
  horPanel.add(panelRight);
  
  scrollPanel.add(horPanel);
  
  mainPanel.add(scrollPanel);
  mainPanel.add(createButton); 
  mainPanel.add(doneButton); 
  
  formPanel.add(mainPanel); 
  
  app.add(formPanel);
  
  DocumentApp.getUi().showModalDialog(app, 'Link creator');
}

/* Function for calling the creation of links between elements */
function doPost(e) {
  if(e.parameter.radio_src == e.parameter.radio_target) {
    DocumentApp.getUi().alert('No link can be created between 2 identical annotations');
  }
  else {
    createLink(e.parameter.radio_src, e.parameter.radio_target);
  }
  displayLinkCreation(); // first idea to solve the pb with doPost closing automatically the window
}

/* Create a link object from the id in the parameters and store it in the cache */
function createLink(idSource, idTarget) {
  var links = JSON.parse(CacheService.getPrivateCache().get("links"));
  if(links[idSource] == undefined) {
    links[idSource] = []; 
  }  
  var target = {};
  target["target"] = idTarget;
  links[idSource].push(target);
  DocumentApp.getUi().alert('Link created (in the cache)');  
  CacheService.getPrivateCache().put('links', JSON.stringify(links), 3600);  
}

/* Save the links into the Document Properties */
function persistLinks(e) {
  var linksInCache = JSON.parse(CacheService.getPrivateCache().get('links'));
  var documentProperties = PropertiesService.getDocumentProperties();
  var linksInDoc = documentProperties.getProperty('LINKS');
  if(linksInDoc == null) {
    documentProperties.setProperty('LINKS', JSON.stringify(linksInCache));
  }
  else {   
    linksInDoc = JSON.parse(documentProperties.getProperty('LINKS'));
    var newLinks = linksInDoc;
    var idsToAdd = Object.keys(linksInCache);
    for (var i = 0; i < idsToAdd.length; i++) {
      if(newLinks[idsToAdd[i]] == undefined) {
        newLinks[idsToAdd[i]] = [];
      }
      /* To avoid creating a new array in the JSON object */
      var nbElem = linksInCache[idsToAdd[i]].length;
      for (var j = 0; j < nbElem; j++) {
        newLinks[idsToAdd[i]].push(linksInCache[idsToAdd[i]][j]);
      }
    }
    documentProperties.setProperty('LINKS', JSON.stringify(newLinks));
  }
  DocumentApp.getUi().alert('Links saved');
  clearLinksInCache(); 
}

/* Create and show a window containing the links */
function showLinks() {
  var app = UiApp.createApplication();
  var mainPanel = app.createVerticalPanel();  
  var scrollPanel = app.createScrollPanel().setPixelSize(500, 300);
  var links = PropertiesService.getDocumentProperties().getProperty('LINKS');
  if(links != null) {
    links = JSON.parse(PropertiesService.getDocumentProperties().getProperty('LINKS'));  
    var ids = Object.keys(links);
    var grid = app.createGrid().resize(ids.length+1,2);
    grid.setBorderWidth(1);
    grid.setText(0, 0, 'Source');
    grid.setText(0, 1, 'Target(s)');
    for (var i = 0; i < ids.length; i++) {
      grid.setText(i+1, 0, ids[i]);
      var linksString = "";
      for (var j = 0; j < links[ids[i]].length; j++) {
        var target = links[ids[i]][j]["target"]
        if(linksString == "") {
          linksString = target;
        }
        else {
          linksString += ", " + target;
        }
      }
      grid.setText(i+1, 1, linksString);
    }
    mainPanel.add(grid); 
    scrollPanel.add(mainPanel);
    app.add(scrollPanel);
  }
  else {
    mainPanel.add(app.createLabel("No links in this document"));
    app.add(mainPanel);
  }
  DocumentApp.getUi().showModalDialog(app, 'Links in the document');  
}

/* Finds the anchored opened comments and return them in a JSON object */
function retrieveAnchoredComments() {
  var fileId = DocumentApp.getActiveDocument().getId();
  /* Default: 20 comments. Acceptable values are 0 to 100, inclusive. */ 
  var comments = Drive.Comments.list(fileId).items;
  
  /* Initialisation of manualAnnotations */
  var annotations_type = JSON.parse(PropertiesService.getDocumentProperties().getProperty('ANNOTATIONS_TYPE')); 
  var types = Object.keys(annotations_type);
  var manualAnnotations = {};
  for (var i = 0; i < types.length; i++) {
    manualAnnotations[types[i]] = [];
  }
  
  /* Filling of manualAnnotations */
  for (var i = 0; i < comments.length; i++) {
    var commentId = comments[i].commentId;
    if(comments[i].status == "open" && comments[i].anchor != undefined) {     
      var comment = JSON.parse(comments[i].content);
      var attributes = Object.keys(comment);
      var type = comment["type"] ;
      var annotation = {};
      for (var j = 1; j < attributes.length; j++) {
        annotation[attributes[j]] = comment[attributes[j]];
      }
      manualAnnotations[type].push(annotation);
      /* The following line isn't working. Updating the content works, but not the status */
      var content = {"status":"resolved"}; 
      Drive.Comments.patch(content, fileId, commentId);
    }
  }
  return manualAnnotations;
}
