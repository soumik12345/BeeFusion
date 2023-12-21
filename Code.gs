function onOpen() {
  var ui = SlidesApp.getUi();
  ui.createMenu('BeeFusion')
      .addItem('Generate an Image from a Prompt', 'generate_image_from_prompt')
      .addItem('Generate an Image from the Slide', 'generate_image_from_slide')
      .addToUi();
}

function generate_image_from_prompt() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('My Custom Sidebar')
    .setWidth(300);
  let ui = SlidesApp.getUi();
  ui.showSidebar(html);
}

function processData(promptValue) {
  generateImage(promptValue);
}

function generate_image_from_slide() {
  var ui = SlidesApp.getUi();
  var response = ui.prompt('Enter index of the Slide', ui.ButtonSet.OK_CANCEL);
  // Process the user's response
  if (response.getSelectedButton() == ui.Button.OK) {
    var slideIndex = Number(response.getResponseText()) - 1;
    var promptFromSlide = extractTextFromSlide(slideIndex);
    // generateImage("A cartoon image of a smiling bee.");
    generateImage(promptFromSlide);
  }
}

function extractTextFromSlide(slideIndex) {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();
  
  var slide = slides[slideIndex];
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

function generateImage(prompt) {
  var presentation = SlidesApp.getActivePresentation();
  var slide = presentation.getSlides()[0];
  var url = "https://api.stability.ai/v1/generation/stable-diffusion-v1-6/text-to-image";
  var payload = JSON.stringify({
    text_prompts: [
      {
        text: prompt,
      },
    ],
    cfg_scale: 7,
    height: 1024,
    width: 1024,
    steps: 30,
    samples: 1,
  });

  var options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'Authorization': 'Bearer sk-QHZZS5CEQpWWG7q6o0xQrHLsIo2hSkDcXnTtLI1fhRU2PXPI',
    },
    payload: payload
  };

  var response = UrlFetchApp.fetch(url, options);
  var data = JSON.parse(response.getContentText());
  Logger.log(data.artifacts[0].seed);
  Logger.log(data.artifacts[0].finishReason);
  var imageBlob = Utilities.newBlob(Utilities.base64Decode(data.artifacts[0].base64), 'image/png', 'sample')
  var image = slide.insertImage(imageBlob);
}