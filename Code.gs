function onOpen() {
  var ui = SlidesApp.getUi();
  ui.createMenu('BeeFusion')
      .addItem('Generate Image', 'generate_image')
      .addItem('WandBot: Verify', 'ask_wandbot_to_verify_text')
      .addItem('WandBot: Correct', 'ask_wandbot_to_correct_text')
      .addToUi();
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

function send_payload_to_wandbot(payload) {
  var url = "https://wandbot.replit.app/query";
  var options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Accept': 'application/json',
    },
    payload: payload
  };
  var response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}

function send_payload_to_openai(payload) {
  var url = "https://api.openai.com/v1/images/generations";
  var openai_api_key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
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
  return JSON.parse(response.getContentText());
}

// -----------------------------------Verify from WandbBot------------------------------------

function get_verification_from_wandbot(text) {
  var question = "Someone told me that \"" + text + "\". Is that true? Please start the answer with a \"Yes\" or \"No\".";
  var data = send_payload_to_wandbot(
    JSON.stringify({
      question: question,
      "application": "BeeFusion",
      "language": "en"
    })
  );
  return data.answer;
}

function ask_wandbot_to_verify_text() {
  const selection = SlidesApp.getActivePresentation().getSelection();
  const selectionType = selection.getSelectionType();
  var ui = SlidesApp.getUi();

  if (selectionType == SlidesApp.SelectionType.PAGE) {
    var currentSlide = selection.getCurrentPage();
    var selectedText = extractTextFromSlide(currentSlide);
    var wandbot_response = get_verification_from_wandbot(selectedText);
    if (wandbot_response.startsWith("Yes")) {
      ui.alert("The information is correct");
    }
    else {
      ui.alert(wandbot_response);
    }
  }
  else if (selectionType == SlidesApp.SelectionType.TEXT) {
    var selectedText = selection.getTextRange().asString();
    var wandbot_response = get_verification_from_wandbot(selectedText);
    if (wandbot_response.startsWith("Yes")) {
      ui.alert("The information is correct");
    }
    else {
      ui.alert(wandbot_response);
    }
  }
  else if (selectionType == SlidesApp.SelectionType.PAGE_ELEMENT) {
    var pageElement = selection.getPageElementRange().getPageElements();
    var selectedText = pageElement[0].asShape().getText().asString();
    var wandbot_response = get_verification_from_wandbot(selectedText);
    if (wandbot_response.startsWith("Yes")) {
      ui.alert("The information is correct");
    }
    else {
      ui.alert(wandbot_response);
    }
  }
  else {
    ui.alert("Nothing is Selected");
  }
}

// -----------------------------------Correct from WandbBot------------------------------------

function ask_wandbot_to_correct_text() {
  const selection = SlidesApp.getActivePresentation().getSelection();
  const selectionType = selection.getSelectionType();
  var ui = SlidesApp.getUi();

  if (selectionType == SlidesApp.SelectionType.PAGE_ELEMENT) {
    var pageElement = selection.getPageElementRange().getPageElements();
    var selectedTextRange = pageElement[0].asShape().getText();
    var selectedText = selectedTextRange.asString()
    var wandbot_response = get_correction_from_wandbot(selectedText);
    if (wandbot_response.startsWith("Yes")) {
      ui.alert("The information is correct");
    }
    else {
      var correct_alternative = wandbot_response.split("A correct alternative could be:")[1];
      correct_alternative = correct_alternative.replace("\"", "");
      correct_alternative = correct_alternative.replace("\"", "");
      selectedTextRange.setText(correct_alternative);
      // ui.alert(correct_alternative);
    }
  }
  else {
    ui.alert("Nothing is Selected");
  }
}

function get_correction_from_wandbot(text) {
  var question = "Original statement: \"" + text + "\"Is the original statement true? Please start the answer with a \"Yes\" or \"No\". What should be the correct alternative to the original statement? Please ensure that the length of the response doesn't exceed that of the original statement by more than 20% starting by the phrase \"A correct alternative could be: \"";
  var data = send_payload_to_wandbot(
    JSON.stringify({
      question: question,
      "application": "BeeFusion",
      "language": "en"
    })
  );
  return data.answer;
}

// ---------------------------------------Generate Image---------------------------------------

function generateImage(text_prompt) {
  var data = send_payload_to_openai(
    JSON.stringify({
      model: "dall-e-3",
      prompt: text_prompt,
      n: 1,
      size: "1024x1024"
    }
  ));
  var decoded_image = Utilities.base64Decode(imageToBase64(data.data[0].url))
  var imageBlob = Utilities.newBlob(decoded_image, 'image/png', 'sample');
  return imageBlob;
}

function generate_image() {
  const selection = SlidesApp.getActivePresentation().getSelection();
  const selectionType = selection.getSelectionType();
  var ui = SlidesApp.getUi();
  var currentSlide = selection.getCurrentPage();
  if (selectionType == SlidesApp.SelectionType.PAGE) {
    var selectedText = extractTextFromSlide(currentSlide);
    var imageBlob = generateImage(selectedText);
    var image = currentSlide.insertImage(imageBlob);
  }
  else if (selectionType == SlidesApp.SelectionType.TEXT) {
    var selectedText = selection.getTextRange().asString();
    var imageBlob = generateImage(selectedText);
    var image = currentSlide.insertImage(imageBlob);
  }
  else if (selectionType == SlidesApp.SelectionType.PAGE_ELEMENT) {
    var pageElement = selection.getPageElementRange().getPageElements();
    var selectedText = pageElement[0].asShape().getText().asString();
    var imageBlob = generateImage(selectedText);
    var image = currentSlide.insertImage(imageBlob);
  }
  else {
    var response = ui.prompt('Enter the Prompt              ', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() == ui.Button.OK) {
      var prompt = response.getResponseText();
      var imageBlob = generateImage(prompt);
      var currentSlide = selection.getCurrentPage();
      var image = currentSlide.insertImage(imageBlob);
    }
  }
}
