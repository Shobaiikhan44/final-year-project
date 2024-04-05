// Add an array to store user-defined special patterns
const specialPatterns = [];

// Function to add a special pattern
function addSpecialPattern() {
  const patternInput = document.getElementById("patternInput");
  const pattern = patternInput.value.trim();

  if (pattern !== "") {
    specialPatterns.push(pattern);
    patternInput.value = "";
    displaySpecialPatterns();
  }
}

// Function to display special patterns
function displaySpecialPatterns() {
  const specialPatternsList = document.getElementById("specialPatternsList");
  specialPatternsList.innerHTML = "";

  specialPatterns.forEach((pattern) => {
    const listItem = document.createElement("li");
    listItem.textContent = pattern;
    specialPatternsList.appendChild(listItem);
  });
}

// Function to check if a position is within any special pattern
function isInSpecialPattern(position) {
  return specialPatterns.some((pattern) => {
    const patternRegex = new RegExp(pattern, "gi");
    return patternRegex.test(position);
  });
}

// Function to be called when the add button is clicked
function onAddSpecialPatternButtonClick() {
  addSpecialPattern();
}

// Attach the onAddSpecialPatternButtonClick function to the button click event
document.getElementById("addSpecialPatternButton").addEventListener("click", onAddSpecialPatternButtonClick);

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



  /////////////

// Function to check and paraphrase text using Language Tool Server
async function checkAndParaphraseText(text) {
    // Define the URL for checking text
    const checkUrl = 'https://api.languagetoolplus.com/v2/check';
  
    // Create a new URLSearchParams object
    const urlencoded = new URLSearchParams();
    urlencoded.append("text", text);
    urlencoded.append("language", "en-US");
  
    // Create headers
    const myHeaders = new Headers();
    myHeaders.append("Content-Type", "application/x-www-form-urlencoded");
    myHeaders.append("Accept", "application/json");
  
    // Create the requestOptions object
    const requestOptions = {
      method: 'POST',
      headers: myHeaders,
      body: urlencoded,
      redirect: 'follow'
    };
  
    try {
      // Send the POST request for checking text
      const response = await fetch(checkUrl, requestOptions);
      const result = await response.json();
  
      // Check if there are matches
      if (result.matches && result.matches.length > 0) {
        // Display suggestions or paraphrased text
        displaySuggestions(result.matches);
      } else {
        console.log("No grammar errors found.");
      }
    } catch (error) {
      console.error('Error checking text:', error);
        }
}
  



// Function to display suggestions or paraphrased text
function displaySuggestions(matches) {
    const suggestionsDiv = document.getElementById("suggestions");
    suggestionsDiv.innerHTML = ""; // Clear previous suggestions
  
    // Initialize an array to store unique paraphrased text
    const uniqueParaphrases = [];
  
    // Loop through matches and display suggestions
    matches.forEach((match) => {
      const suggestionElement = document.createElement("div");
      suggestionElement.classList.add("suggestion");
  
      // Display the message
      const messageElement = document.createElement("p");
      messageElement.textContent = match.message;
      suggestionElement.appendChild(messageElement);
  
      // Display paraphrased text if available and not already displayed
      if (match.replacements && match.replacements.length > 0) {
        const paraphrasedText = match.replacements[0].value;
        if (!uniqueParaphrases.includes(paraphrasedText)) {
          uniqueParaphrases.push(paraphrasedText);
  
          const paraphrasedTextElement = document.createElement("p");
          paraphrasedTextElement.textContent = `-Paraphrased Word: ${paraphrasedText}`;
          suggestionElement.appendChild(paraphrasedTextElement);
        }
      }
  
      suggestionsDiv.appendChild(suggestionElement);
    });
  
    // Display the complete text with corrections
    const correctedText = matches.reduce((text, match) => {
      return text.substring(0, match.offset) + (match.replacements.length > 0 ? match.replacements[0].value : match.context.text) + text.substring(match.offset + match.length);
    }, matches[0].context.text);
  
    const correctedTextElement = document.createElement("div");
    correctedTextElement.classList.add("corrected-text");
    correctedTextElement.textContent = `-Corrected Text: ${correctedText}`;
    suggestionsDiv.appendChild(correctedTextElement);
  }
  
  // Function to insert the paraphrased text into the document
  async function insertParaphrasedText(paraphrasedText) {
    try {
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.insertText(paraphrasedText, Word.InsertLocation.replace);
  
        await context.sync();
      });
    } catch (error) {
      console.error("Error inserting paraphrased text:", error);
    }
  }
  

  // Add event listener to the "Check and Paraphrase" button
  document.getElementById("checkParaphraseButton").addEventListener("click", async () => {
    try {
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.load("text");
  
        await context.sync();
  
        const textToCheck = range.text;
  
        if (textToCheck) {
          // Call the function to check and paraphrase the selected text
          checkAndParaphraseText(textToCheck);
        } else {
          console.log("No text selected.");
        }
      });
    } catch (error) {
      console.error("Error:", error);
    }
  });
  

