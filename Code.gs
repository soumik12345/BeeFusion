function onOpen() {
  var ui = SlidesApp.getUi();
  ui.createMenu('BeeFusion')
      .addSubMenu(ui.createMenu('Generate an Image from a Prompt')
        .addItem('Style: None', 'generate_image_from_prompt_no_style')
        .addItem('Style: GPT-4 Enhanced', 'generate_image_from_prompt_gpt4_enhanced')
        .addItem('Style: Default Sharp', 'generate_image_from_prompt_default_sharp'))
      .addSubMenu(ui.createMenu('Generate an Image from Slide')
        .addItem('Style: None', 'generate_image_from_slide_no_style')
        .addItem('Style: GPT-4 Enhanced', 'generate_image_from_slide_gpt4_enhanced')
        .addItem('Style: Default Sharp', 'generate_image_from_slide_default_sharp'))
      .addToUi();
}

// -----------------------------Generate Image From a Prompt-----------------------------

function generate_image_from_prompt_no_style() {
  var ui = SlidesApp.getUi();
  var response = ui.prompt('Enter the Prompt              ', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var prompt = response.getResponseText();
    var text_prompts = [
      {
        text: prompt, weight: 1.0,
      },
    ]
    var imageBlob = generateImage(text_prompts);
    var currentSlide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
    var image = currentSlide.insertImage(imageBlob);
  }
}

function generate_image_from_prompt_gpt4_enhanced() {
  var ui = SlidesApp.getUi();
  var response = ui.prompt('Enter the Prompt              ', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var prompt = response.getResponseText();
    var enhancedPrompt = upsamplePrompt(prompt);
    var text_prompts = [
      {
        text: enhancedPrompt, weight: 1.0,
      },
    ]
    var imageBlob = generateImage(text_prompts);
    var currentSlide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
    var image = currentSlide.insertImage(imageBlob);
  }
}

function generate_image_from_prompt_default_sharp() {
  var ui = SlidesApp.getUi();
  var response = ui.prompt('Enter the Prompt              ', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var prompt = response.getResponseText();
    var text_prompts = [
      {
        text: "cinematic still " + prompt + " . emotional, harmonious, vignette, 4k epic detailed, shot on kodak, 35mm photo, sharp focus, high budget, cinemascope, moody, epic, gorgeous, film grain, grainy", weight: 1.0,
      },
      {
        text: "anime, cartoon, graphic, (blur, blurry, bokeh), text, painting, crayon, graphite, abstract, glitch, deformed, mutated, ugly, disfigured", weight: -1.0,
      },
    ]
    var imageBlob = generateImage(text_prompts);
    var currentSlide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
    var image = currentSlide.insertImage(imageBlob);
  }
}

// -----------------------------Generate Image From a Slide-----------------------------

function generate_image_from_slide_no_style() {
  var currentSlide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
  var promptFromSlide = extractTextFromSlide(currentSlide);
  var text_prompts = [
    {
      text: promptFromSlide, weight: 1.0,
    },
  ]
  var imageBlob = generateImage(text_prompts);
  var image = currentSlide.insertImage(imageBlob);
}

function generate_image_from_slide_gpt4_enhanced() {
  var currentSlide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
  var promptFromSlide = extractTextFromSlide(currentSlide);
  var enhancedPromptFromSlide = upsamplePrompt(promptFromSlide);
  var text_prompts = [
    {
      text: enhancedPromptFromSlide, weight: 1.0,
    },
  ]
  var imageBlob = generateImage(text_prompts);
  var image = currentSlide.insertImage(imageBlob);
}

function generate_image_from_slide_default_sharp() {
  var currentSlide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
  var promptFromSlide = extractTextFromSlide(currentSlide);
  var text_prompts = [
    {
      text: "cinematic still " + promptFromSlide + " . emotional, harmonious, vignette, 4k epic detailed, shot on kodak, 35mm photo, sharp focus, high budget, cinemascope, moody, epic, gorgeous, film grain, grainy", weight: 1.0,
    },
    {
      text: "anime, cartoon, graphic, (blur, blurry, bokeh), text, painting, crayon, graphite, abstract, glitch, deformed, mutated, ugly, disfigured", weight: -1.0,
    },
  ];
  var imageBlob = generateImage(text_prompts);
  var image = currentSlide.insertImage(imageBlob);
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

// ---------------------------------------Upsample Prompt---------------------------------------

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

// ---------------------------------------Generate Image---------------------------------------

function generateImage(text_prompts) {
  var url = "https://api.stability.ai/v1/generation/stable-diffusion-xl-1024-v1-0/text-to-image";
  var payload = JSON.stringify({
    text_prompts: text_prompts,
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
