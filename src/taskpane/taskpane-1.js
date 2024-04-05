/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// Office.onReady((info) => {
//   if (info.host === Office.HostType.Word) {
//     document.getElementById("sideload-msg").style.display = "none";
//     document.getElementById("app-body").style.display = "flex";
//     document.getElementById("run").onclick = run;
//   }
// });

// export async function run() {
//   return Word.run(async (context) => {
//     /**
//      * Insert your Word code here
//      */

//     // insert a paragraph at the end of the document.
//     const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

//     // change the paragraph color to blue.
//     paragraph.font.color = "blue";

//     await context.sync();
//   });
// }




// Manually Adding 


(function () {
  "use strict";

  const checkGrammarButton = document.getElementById("checkGrammarButton");
  const clearSuggestionsButton = document.getElementById("clearSuggestionsButton");
  const suggestionsDiv = document.getElementById("suggestions");

  // Define a function to send text to the Language Tool Server for grammar checking
  async function sendToLanguageTool(text) {

      clearSuggestionsButton.style.display = "none";
      // Disable the "Check Grammar" button and update its text
      checkGrammarButton.disabled = true;
      checkGrammarButton.textContent = "Checking...";

      // Define the URL
      const requestUrl = 'https://api.languagetoolplus.com/v2/check';

      // Create a new URLSearchParams object
      var urlencoded = new URLSearchParams();
      urlencoded.append("text", text);
      urlencoded.append("language", "en-US");
      urlencoded.append("enabledOnly", "false");

      // Create headers
      var myHeaders = new Headers();
      myHeaders.append("Content-Type", "application/x-www-form-urlencoded");
      myHeaders.append("Accept", "application/json");

      // Create the requestOptions object
      var requestOptions = {
          method: 'POST',
          headers: myHeaders,
          body: urlencoded,
          redirect: 'follow'
      };

      // Send the POST request
      fetch(requestUrl, requestOptions)
          .then(response => response.json())
          .then(result => {
              // Process the result here
              console.log(result);
              displaySuggestions(result);
          })
          .catch(error => console.log('error', error))
          .finally(() => {
              // Re-enable the "Check Grammar" button and reset its text
              checkGrammarButton.disabled = false;
              checkGrammarButton.textContent = "Check Grammar";
          });

  }


  // Define a function to display grammar suggestions
  function displaySuggestions(suggestions) {

      // Clear previous suggestions
      suggestionsDiv.innerHTML = "";

      // Show the "Clear Suggestions" button
      clearSuggestionsButton.style.display = "block";

      // Check if there are no matches
      if (suggestions.matches.length === 0) {
          const noErrorsMessage = document.createElement("p");
          noErrorsMessage.textContent = "No grammar errors found.";
          suggestionsDiv.appendChild(noErrorsMessage);
          return; // Exit the function, as there are no suggestions to display
      }

      // Loop through suggestions and display them
      suggestions.matches.forEach((suggestion) => {
          const suggestionElement = document.createElement("div");
          suggestionElement.classList.add("suggestion");

          // Display the message
          const messageElement = document.createElement("p");
          messageElement.textContent = suggestion.message;
          suggestionElement.appendChild(messageElement);

          // Display replacements if available
          if (suggestion.replacements && suggestion.replacements.length > 0) {
              // Message about possible replacements
              const replacementsMessage = document.createElement("p");
              replacementsMessage.textContent = "Possible replacements:";
              replacementsMessage.classList.add("replacements-message");
              suggestionElement.appendChild(replacementsMessage);

              // List of replacements
              const replacementsElement = document.createElement("ul");
              replacementsElement.classList.add("replacements");

              suggestion.replacements.forEach((replacement) => {
                  const replacementItem = document.createElement("li");
                  replacementItem.textContent = replacement.value;
                  replacementsElement.appendChild(replacementItem);
              });

              suggestionElement.appendChild(replacementsElement);
          }

          suggestionsDiv.appendChild(suggestionElement);
      });
  }

  // Add event listener to the "Check Grammar" button
  document.getElementById("checkGrammarButton").addEventListener("click", async () => {
      try {
          await Word.run(async (context) => {
              const range = context.document.getSelection();
              range.load("text");

              await context.sync();

              const textToCheck = range.text;

              if (textToCheck) {
                  // Send the selected text to the Language Tool Server
                  sendToLanguageTool(textToCheck);
              } else {
                  console.log("No text selected.");
              }
          });
      } catch (error) {
          console.error("Error:", error);
      }
  });

  // Add event listener to the "Clear Suggestions" button
  clearSuggestionsButton.addEventListener("click", () => {
      suggestionsDiv.innerHTML = "";
      clearSuggestionsButton.style.display = "none";
  });

})();
