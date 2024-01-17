function onOpen() {
  var ui = SlidesApp.getUi();
  ui.createMenu('BeeFusion')
      .addItem('Generate Image', 'generate_image')
      .addItem('Verify text from WandBot', 'ask_wandbot')
      .addToUi();
}

// -----------------------------------Generate Image-----------------------------------

function generate_image() {
  const selection = SlidesApp.getActivePresentation().getSelection();
  const selectionType = selection.getSelectionType();
  var ui = SlidesApp.getUi();
  if (selectionType == SlidesApp.SelectionType.PAGE) {
    var currentSlide = selection.getCurrentPage();
    var selectedText = extractTextFromSlide(currentSlide);
    ui.alert(selectedText);
    var imageBlob = generateImage(selectedText);
    var image = currentSlide.insertImage(imageBlob);
  }
  else if (selectionType == SlidesApp.SelectionType.TEXT) {
    var currentSlide = selection.getCurrentPage();
    var selectedText = selection.getTextRange().asString();
    ui.alert(selectedText);
    var imageBlob = generateImage(selectedText);
    var image = currentSlide.insertImage(imageBlob);
  }
  else {
    var response = ui.prompt('Enter the Prompt              ', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() == ui.Button.OK) {
      var prompt = response.getResponseText();
      ui.alert(prompt);
      var imageBlob = generateImage(prompt);
      var currentSlide = selection.getCurrentPage();
      var image = currentSlide.insertImage(imageBlob);
    }
  }
}

// -----------------------------------Verify from WandbBot------------------------------------

function ask_wandbot() {
  const selection = SlidesApp.getActivePresentation().getSelection();
  const selectionType = selection.getSelectionType();
  var ui = SlidesApp.getUi();

  if (selectionType == SlidesApp.SelectionType.PAGE) {
    var currentSlide = selection.getCurrentPage();
    var selectedText = extractTextFromSlide(currentSlide);
    ui.alert(selectedText);
  }
  else if (selectionType == SlidesApp.SelectionType.TEXT) {
    var selectedText = selection.getTextRange().asString();
  }
  else if (selectionType == SlidesApp.SelectionType.NONE) {
    ui.alert("Nothing is Selected");
  }
}

// -------------------------------------------Utils-------------------------------------------

function extractTextFromSlide(slide) {
  var textElements = slide.getPageElements().filter(function(element) {
    return element.getPageElementType() === SlidesApp.PageElementType.SHAPE;
  });

  var allText = '';
  textElements.forEach(function(element) {
    var textShape = element.asShape();
    if (textShape.getText()) {
      allText += textShape.getText().asString() + '\n';
    }
  });

  return allText;
}

function imageToBase64(url) {
  try {
    // Fetch the image
    var response = UrlFetchApp.fetch(url);
    var blob = response.getBlob();

    // Convert the image to a base64 string
    var base64String = Utilities.base64Encode(blob.getBytes());

    return base64String;
  } catch (e) {
    Logger.log("Error in fetching or converting the image: " + e.toString());
    return null;
  }
}

// ---------------------------------------Generate Image---------------------------------------

function generateImage(text_prompt) {
  // var ui = SlidesApp.getUi();
  // ui.alert(text_prompt);
  var url = "https://api.openai.com/v1/images/generations";
  var payload = JSON.stringify({
    model: "dall-e-3",
    prompt: text_prompt,
    n: 1,
    size: "1024x1024"
  });

  var openai_api_key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY')
  var options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'Authorization': 'Bearer ' + openai_api_key,
    },
    payload: payload
  };

  var response = UrlFetchApp.fetch(url, options);
  var data = JSON.parse(response.getContentText());
  var decoded_image = Utilities.base64Decode(imageToBase64(data.data[0].url))
  var imageBlob = Utilities.newBlob(decoded_image, 'image/png', 'sample');
  return imageBlob;
}
