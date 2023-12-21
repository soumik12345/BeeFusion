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
  var imageBlob = generateImage(promptValue);
  var currentSlide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
  var image = currentSlide.insertImage(imageBlob);
}

function generate_image_from_slide() {
  var ui = SlidesApp.getUi();
  var response = ui.prompt('Enter index of the Slide', ui.ButtonSet.OK_CANCEL);
  // Process the user's response
  if (response.getSelectedButton() == ui.Button.OK) {
    var slideIndex = Number(response.getResponseText()) - 1;
    var promptFromSlide = extractTextFromSlide(slideIndex);
    var imageBlob = generateImage(promptFromSlide);
    var outputSlide = SlidesApp.getActivePresentation().getSlides()[slideIndex];
    var image = outputSlide.insertImage(imageBlob);
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
  var url = "https://api.stability.ai/v1/generation/stable-diffusion-xl-1024-v1-0/text-to-image";
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
      'Authorization': 'Bearer ${api-key}',
    },
    payload: payload
  };

  var response = UrlFetchApp.fetch(url, options);
  var data = JSON.parse(response.getContentText());
  Logger.log(data.artifacts[0].seed);
  Logger.log(data.artifacts[0].finishReason);
  var imageBlob = Utilities.newBlob(Utilities.base64Decode(data.artifacts[0].base64), 'image/png', 'sample');
  return imageBlob;
}
