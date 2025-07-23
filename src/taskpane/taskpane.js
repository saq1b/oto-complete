/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("run").onclick = run;
    // Add event listener for typing
    document.addEventListener("keyup", async (event) => {
      if (event.key.length === 1 || event.key === "Backspace") { // Trigger only for valid typing keys
        await getLastWordAndSuggestions();
      }
    });
  }
});

export async function run() {
  return Word.run(async (context) => {
    console.log("Run Clicked");
    const typedWords = new Set();

    async function trackWords(context) {
      const body = context.document.body;
      body.load("text"); // Load the text property of the body
      await context.sync(); // Sync the context to retrieve the loaded data

      const words = body.text.match(/\b\w+\b/g); // Access the text property
      if (words) {
        words.forEach(word => typedWords.add(word));
      }
    }

    await trackWords(context);
    console.log(typedWords);

    function getSuggestions(prefix) {
      const lowerPrefix = prefix.toLowerCase();
      const suggestions = Array.from(typedWords).filter(word =>
        word.toLowerCase().startsWith(lowerPrefix)
      );
      return suggestions.slice(0, 5); // Show top 5 suggestions
    }

    // New function to display suggestions in a dropdown at cursor position
    async function showSuggestionsDropdown(suggestions) {
      // Use a fixed position for the dropdown
      const position = { left: 100, top: 200 }; // Example fixed position

      createDropdown(suggestions, position);
    }

    function createDropdown(suggestions, position) {
      // Remove existing dropdown if any
      const existingDropdown = document.getElementById("suggestions-dropdown");
      if (existingDropdown) {
        existingDropdown.remove();
      }

      // Create a new dropdown
      const dropdown = document.createElement("div");
      dropdown.id = "suggestions-dropdown";
      dropdown.style.position = "absolute";
      dropdown.style.left = `${position.left}px`;
      dropdown.style.top = `${position.top}px`;
      dropdown.style.border = "1px solid #ccc";
      dropdown.style.backgroundColor = "#fff";
      dropdown.style.zIndex = "1000";
      dropdown.style.padding = "5px";
      dropdown.style.boxShadow = "0px 4px 6px rgba(0, 0, 0, 0.1)";

      // Add suggestions to the dropdown
      suggestions.forEach(suggestion => {
        const item = document.createElement("div");
        item.textContent = suggestion;
        item.style.padding = "5px";
        item.style.cursor = "pointer";
        item.addEventListener("click", async () => {
          Word.run(async (context) => {
            const range = context.document.getSelection();
            range.insertText(suggestion, Word.InsertLocation.replace);
            await context.sync();
          });
          dropdown.remove(); // Remove dropdown after selection
        });
        dropdown.appendChild(item);
      });

      document.body.appendChild(dropdown);
    }

    // Example usage of the dropdown function
    /*const suggestions = getSuggestions("he");
    console.log(suggestions); 
    await showSuggestionsDropdown(suggestions);*/

    // // Insert a paragraph at the end of the document.
    // const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // // Change the paragraph color to blue.
    // paragraph.font.color = "blue";
    getLastWordAndSuggestions();
    // Sync the context to apply changes

    await context.sync();
  
    async function getLastWordAndSuggestions() {
      await Word.run(async (context) => {
        console.log("Getting last word and suggestions");
        // Load the paragraphs in the document
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("items");
        await context.sync();
    
        // Get the last paragraph (or the one being edited)
        const lastParagraph = paragraphs.items[paragraphs.items.length - 1];
        lastParagraph.load("text");
        await context.sync();
    
        // Extract the last word from the paragraph text
        const text = lastParagraph.text;
        console.log("Current Paragraph:", text);
    
        const lastWord = text.split(/\s+/).pop(); // Get the last word typed
        console.log("Last Word:", lastWord);
    
        // Get suggestions based on the last word
        const suggestions = getSuggestions(lastWord);
        console.log("Suggestions:", suggestions);
    
        // Display the suggestions in a dropdown
        showDropdownInDocument(suggestions);
      });
    }
    
    function showDropdownInDocument(suggestions) {
      // Example fixed position for the dropdown
      const position = { left: 100, top: 200 };
    
      // Remove existing dropdown if any
      const existingDropdown = document.getElementById("suggestions-dropdown");
      if (existingDropdown) {
        existingDropdown.remove();
      }
    
      // Create a new dropdown
      const dropdown = document.createElement("div");
      dropdown.id = "suggestions-dropdown";
      dropdown.style.position = "absolute";
      dropdown.style.left = `${position.left}px`;
      dropdown.style.top = `${position.top}px`;
      dropdown.style.border = "1px solid #ccc";
      dropdown.style.backgroundColor = "#fff";
      dropdown.style.zIndex = "1000";
      dropdown.style.padding = "5px";
      dropdown.style.boxShadow = "0px 4px 6px rgba(0, 0, 0, 0.1)";
    
      // Add suggestions to the dropdown
      suggestions.forEach(suggestion => {
        const item = document.createElement("div");
        item.textContent = suggestion;
        item.style.padding = "5px";
        item.style.cursor = "pointer";
        item.addEventListener("click", async () => {
          await insertSuggestion(suggestion);
          dropdown.remove(); // Remove dropdown after selection
        });
        dropdown.appendChild(item);
      });
    
      document.body.appendChild(dropdown);
    }
    
    async function insertSuggestion(word) {
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.insertText(word + " ", Word.InsertLocation.replace);
        await context.sync();
      });
    }
  });

}