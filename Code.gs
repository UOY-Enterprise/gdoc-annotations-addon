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
  }
}

/* Adds a type to the selected text */
function save(type, color) {   
  var selectedElements = getSelectedText();
  addToCache(type, selectedElements.text, selectedElements.elements);
  updateView(selectedElements.elements, color);
  insertComment(DocumentApp.getActiveDocument().getId(), selectedElements.text, type);
}

/* Adds the selected text to the cache */
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
  // NB : returns the last matching element
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
  for (var i = 0; i < xml_content.length; i++) {
    annotations[xml_content[i].name] = [];
    annotations_type[xml_content[i].name] = {"color" : xml_content[i].color, "attributes" : new Object(xml_content[i].attributes)};
  }  
  CacheService.getPrivateCache().put('annotations', JSON.stringify(annotations), 3600);   
  PropertiesService.getDocumentProperties().setProperty('ANNOTATIONS_TYPE', JSON.stringify(annotations_type));
}

/* Clears the cache */
function clearCache() {
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

/* Saves the annotations in the document properties */
function saveAnnotationsInDoc() {
  var annotationsInCache = JSON.parse(CacheService.getPrivateCache().get('annotations'));
  var documentProperties = PropertiesService.getDocumentProperties();
  var annotationsInDoc = documentProperties.getProperty('ANNOTATIONS');
  var annotations_type = JSON.parse(documentProperties.getProperty('ANNOTATIONS_TYPE'));
  if(annotationsInDoc == null) {
    documentProperties.setProperty('ANNOTATIONS', JSON.stringify(annotationsInCache));
  }
  else {
    annotationsInDoc = JSON.parse(documentProperties.getProperty('ANNOTATIONS'));
    var newAnnotations = annotationsInDoc;
    var keys = Object.keys(annotationsInDoc);
    for(var k in keys) {
      var type = keys[k];
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
    }
    documentProperties.setProperty('ANNOTATIONS', JSON.stringify(newAnnotations));
  }
  clearCache();
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
  DocumentApp.getUi().alert('dans les prop du document : ' + documentProperties.getProperty('ANNOTATIONS'));
  //DocumentApp.getUi().alert('dans les prop du document type : ' + documentProperties.getProperty('ANNOTATIONS_TYPE'));  
}

/* Remove the annotations stored in the document properties */
function clearAnnotationsInDoc() {
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.deleteProperty('ANNOTATIONS');
  documentProperties.deleteProperty('ANNOTATIONS_TYPE'); 
  documentProperties.deleteProperty('ID'); 
  // A reload of the page is needed after this function, otherwise the types are empty
}

/* Parses the XML file and customise the application with the new types */
function loadXML(e){
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
  
  newSidebar += '<tr><button onclick="google.script.run.getAnnotationsInDoc()">What\'s in the doc</button>';
  newSidebar += '<button onclick="google.script.run.saveAnnotationsInDoc()" class="action">Save</button></tr>';
  
  newSidebar += '<tr><button onclick="google.script.run.clearAnnotationsInDoc()">Clear doc prop</button>';
  newSidebar += '<button onclick="google.script.run.generateSpreadsheet()" class="action">Generate Spreadsheet</button></tr>';
  
  newSidebar += '<tr><button onclick="google.script.run.clearCache()">Clear cache</button></tr>';
  // newSidebar += '<button onclick="google.script.run.customise()">customise</button>';
  
  newSidebar += '</table>';
  
  newSidebar +='</div> </div> </div>';
  newSidebar +='<script> function save(button) {google.script.run.save(button.name, button.value);}</script>';
  var htmlOutput = HtmlService.createHtmlOutput(newSidebar);
  htmlOutput.setWidth(250).setHeight(300);
  DocumentApp.getUi().showSidebar(htmlOutput);
}

/* Window for creating links between elements */
function createLink() {
  var app = UiApp.createApplication();    
  var mainPanel = app.createVerticalPanel();
  
  var panelRadio = app.createVerticalPanel(); 
  var validateButton = app.createButton('validate');
  var textRadio1 = 'texte du premier radio button';
  var textRadio2 = 'texte du deuxieme radio button';
  var textRadio3 = 'texte du troisieme radio button';
  
  var radio1 = app.createRadioButton('radio_ob', textRadio1);
  var radio2 = app.createRadioButton('radio_ob', textRadio2);
  var radio3 = app.createRadioButton('radio_ob', textRadio3);
  
  var panelCheck = app.createVerticalPanel(); 
  var textCheck1 = 'diagnosis1';
  var textCheck2 = 'diagnosis2';
  var textCheck3 = 'diagnosis3';
  
  var check1 = app.createCheckBox(textCheck1);
  var check2 = app.createCheckBox(textCheck2);
  var check3 = app.createCheckBox(textCheck3);
  
 // var textbox_to_test = app.createTextBox().setName('nom1');
 // mainPanel.add(textbox_to_test); 
  panelRadio.add(radio1);   panelRadio.add(radio2);   panelRadio.add(radio3);
  panelCheck.add(check1); panelCheck.add(check2); panelCheck.add(check3); 
  
  var clickHandler = app.createServerHandler('validateLink');
  validateButton.addClickHandler(clickHandler);
  clickHandler.addCallbackElement(mainPanel);
  var horPanel = app.createHorizontalPanel();
  horPanel.add(panelRadio);
  horPanel.add(panelCheck);
  mainPanel.add(horPanel);
  mainPanel.add(validateButton); 
  app.add(mainPanel);
  
  DocumentApp.getUi().showModalDialog(app, 'Link creator');
}

/* Function for creating links between elements */
function validateLink(e) {
  /* TODO: continue this function */
  DocumentApp.getUi().alert(e.parameter.nom1);
}
