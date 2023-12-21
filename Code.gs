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
  var response = ui.prompt('Enter the slide number that you’d like to generate the image for', ui.ButtonSet.OK_CANCEL);
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

function upsamplePrompt(prompt) {
  var url = "https://api.openai.com/v1/chat/completions";
  messages = [
    {
      "role": "system",
      "content": `You are part of a team of bots that creates images. You work with an assistant bot that will draw anything you say in square brackets. For example, outputting "a beautiful morning in the woods with the sun peaking through the trees" will trigger your partner bot to output an image of a forest morning, as described. You will be prompted by people looking to create detailed, amazing images. The way to accomplish this is to take their short prompts and make them extremely detailed and descriptive.
      There are a few rules to follow:
      - You will only ever output a single image description per user request.
      - Sometimes the user will request that you modify previous captions. In this case, you should refer to your previous conversations with the user and make the modifications requested.
      - When modifications are requested, you should not simply make the description longer. You should refactor the entire description to integrate the suggestions.
      - Other times the user will not want modifications, but instead want a new image. In this case, you should ignore your previous conversation with the user."
      - Image descriptions must be between 15-77 words. Extra words will be ignored.
      `,
    },
    {
      "role": "user",
      "content": "Create an imaginative image descriptive caption or modify an earlier caption for the user input: 'a man holding a sword'",
    },
    {
      "role": "assistant",
      "content": "a pale figure with long white hair stands in the center of a dark forest, holding a sword high above his head. the blade glows with a blue light , casting a soft glow on the trees and bushes surrounding him.",
    },
    {
      "role": "user",
      "content": "Create an imaginative image descriptive caption or modify an earlier caption for the user input : 'make the light red'",
    },
    {
      "role": "assistant",
      "content": "a pale figure with long white hair stands in the center of a dark forest, holding a sword high above his head. the blade glows with a red light, casting a warm glow on the trees and bushes surrounding him.",
    },
    {
      "role": "user",
      "content": "Create an imaginative image descriptive caption or modify an earlier caption for the user input : 'draw a frog playing dominoes'",
    },
    {
      "role": "assistant",
      "content": "a frog sits on a worn table playing a game of dominoes with an elderly raccoon. the table is covered in a green cloth, and the frog is wearing a jacket and a pair of jeans. The scene is set in a forest, with a large tree in the background.",
    },
    {
      "role": "user",
      "content": "Create an imaginative image descriptive caption or modify an earlier caption for the user input : '" + prompt + "'",
    },
  ];
  var payload = JSON.stringify({
    "model": "gpt-4",
    "messages": messages,
  });

  var options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'Authorization': 'Bearer {openai-api}',
    },
    payload: payload
  };

  var response = UrlFetchApp.fetch(url, options);
  var data = JSON.parse(response.getContentText());
  return data.choices[0].message.content;
}

function generateImage(prompt) {
  prompt = upsamplePrompt(prompt)
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
      'Authorization': 'Bearer {stability-api}',
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
