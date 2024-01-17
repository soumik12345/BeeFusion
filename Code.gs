function onOpen() {
  var ui = SlidesApp.getUi();
  ui.createMenu('BeeFusion')
      .addSubMenu(ui.createMenu('Generate Image')
        .addItem('Generate from Prompt', 'generate_image_from_prompt')
        .addItem('Generate from Slide', 'generate_image_from_slide'))
      .addItem('Verify text from WandBot', 'get_verification_from_wandbot')
      .addToUi();
}

// -----------------------------Generate Image From a Prompt-----------------------------

function generate_image_from_prompt() {
  var ui = SlidesApp.getUi();
  var response = ui.prompt('Enter the Prompt              ', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var prompt = response.getResponseText();
    var imageBlob = generateImage(prompt);
    var currentSlide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
    var image = currentSlide.insertImage(imageBlob);
  }
}

// -----------------------------Generate Image From a Slide-----------------------------

function generate_image_from_slide() {
  var currentSlide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
  var promptFromSlide = extractTextFromSlide(currentSlide);
  var imageBlob = generateImage(promptFromSlide);
  var image = currentSlide.insertImage(imageBlob);
}

// -----------------------------------Verify from WandbBot------------------------------------

function get_verification_from_wandbot() {
  const selection = SlidesApp.getActivePresentation().getSelection();
  const selectionType = selection.getSelectionType();
  var ui = SlidesApp.getUi();

  if (selectionType == SlidesApp.SelectionType.PAGE) {
    var currentSlide = selection.getCurrentPage();
    var promptFromSlide = extractTextFromSlide(currentSlide);
    ui.alert(promptFromSlide);
  }
  else {
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

  var options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'Authorization': 'Bearer ' + PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'),
    },
    payload: payload
  };

  var response = UrlFetchApp.fetch(url, options);
  var data = JSON.parse(response.getContentText());
  var imageBlob = Utilities.newBlob(Utilities.base64Decode(imageToBase64(data.data[0].url)), 'image/png', 'sample');
  return imageBlob;
}
