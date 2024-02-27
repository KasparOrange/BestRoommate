/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";
    // document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    // /**
    //  * Insert your Word code here
    //  */

    // // insert a paragraph at the end of the document.
    // const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // // change the paragraph color to blue.
    // paragraph.font.color = "blue";

    await context.sync();
  });
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    createFakeNavPane();
    document.getElementById("test-button-1").onclick = findRange;
    document.getElementById("clear-log-button").onclick = clearLogs;
    document.getElementById("test-button-2").onclick = insertTextAtSelection;
    document.getElementById("test-button-3").onclick = update;
  }
});

function update() {
  Word.run(async (context) => {
    createFakeNavPane()
    context.sync();
  });
}

function insertTextAtSelection() {
  Word.run(async (context) => {
    var range = context.document.getSelection();
    range.insertText("Hello World", "End");
    return context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });

}

function clearLogs() {
  const logs = document.getElementById("customLogs");
  logs.innerHTML = "";
}

function log(text) {
  console.log(text);
  const logs = document.getElementById("customLogs");
  const li = document.createElement("li");
  li.textContent = text;
  logs.appendChild(li);
}

async function findRange() {
  await Word.run(async (context) => {
    const logs = document.getElementById("customLogs");
    // logs.innerHTML = "";
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text, style");

    await context.sync();

    log(`MYX Number of paragraphs: ${paragraphs.items.length}`)

    logs.innerHTML = `Number of paragraphs: ${paragraphs.items.length}`;
    let headingsList = new Set();
    paragraphs.items.forEach(async (paragraph) => {
      // if (paragraph.text.contains("Diplomarbeit")) {
      //   log("MYX findRange 4");
      //   // paragraph.load("text");


      /** @type {Word.Range} */
      let range = paragraph.getRange(Word.RangeLocation.whole);

      range.load("text");
      await context.sync();
      log(range.text);
      // log("MYX findRange 4");
      //   log(paragraph.text);
      // }

      // if (paragraph.style.startsWith("heading")) {
      //   headingsList.add(paragraph);
      //   // console.log("MYX findRange 3");
      //   paragraph.load("text");
      //   await context.sync();
      //   console.log(paragraph.text);
      // }
    });
    // console.log(`MYX Number of headings: ${headingsList.size}`);

  }).catch((error) => console.error(error));
}

// async function createFakeNavPane() {
//   let headingCounter = 0; // Counter to generate unique IDs for headings

//   await Word.run(async (context) => {
//     const paragraphs = context.document.body.paragraphs;
//     paragraphs.load("text, style");
//     await context.sync();

//     let headingLevels = [];
//     const headingsList = document.getElementById("headingsList");
//     headingsList.innerHTML = ""; // Clear existing items

//     paragraphs.items.forEach((paragraph, index) => {
//       if (paragraph.style.startsWith("Heading")) {
//         const level = parseInt(paragraph.style.replace("Heading ", "")) - 1;
//         headingLevels[level] = (headingLevels[level] || 0) + 1; // Increment count for current level

//         // Reset counts for all lower levels
//         for (let i = level + 1; i < headingLevels.length; i++) {
//           headingLevels[i] = 0;
//         }

//         // Construct the heading number (e.g., "8.3.2")
//         const headingNumber = headingLevels.slice(0, level + 1).join(".");
//         const headingID = `heading-${headingCounter++}`; // Generate a unique ID

//         // Create the list item
//         const listItem = document.createElement("li");
//         listItem.textContent = `${headingNumber} ${paragraph.text}`;
//         listItem.setAttribute("id", headingID); // Set the ID on the listItem
//         listItem.style.cursor = "pointer"; // Change cursor on hover
//         headingsList.appendChild(listItem);

//         // Attach event listener for navigation
//         listItem.addEventListener("click", () => navigateToHeading(context, index));
//       }
//     });

//     await context.sync();
//   }).catch(error => console.error("Error:", error));
// }

async function navigateToHeading(context, paragraphIndex, event) {
  await context.sync().then(async () => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('text');
    await context.sync();

    if (event.ctrlKey || event.metaKey) {
      // If Ctrl or Cmd key is pressed, add a checkmark to the heading
      // console.log("MYX Ctrl or Cmd key is pressed");
      const targetParagraph = paragraphs.items[paragraphIndex];
      // addSignToHeading(targetParagraph); // Add a checkmark to the heading
      targetParagraph.insertText("✓", Word.InsertLocation.end);
      await context.sync();
      createFakeNavPane(); // Refresh the nav pane to reflect the updates
      return;
    }
    // Scroll to the specific paragraph
    const targetParagraph = paragraphs.items[paragraphIndex];
    targetParagraph.select();
    await context.sync();
  });
}

async function createFakeNavPane() {
  let headingCounter = 0; // Counter to generate unique IDs for headings

  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text, style");

    await context.sync();

    let headingLevels = [];

    const headingsList = document.getElementById("headingsList");
    headingsList.innerHTML = ""; // Clear existing items

    // let headingLevel = "";

    // let bla = [];

    // let headingLevels = [];

    paragraphs.items.forEach(async (paragraph, index) => {
      if (paragraph.style.startsWith("heading")) {

        // // remove heading from style
        // paragraph.load("text, style");
        // await context.sync();
        const level = parseInt(paragraph.style.replace("heading ", "")) - 1;
        headingLevels[level] = (headingLevels[level] || 0) + 1; // Increment count for current level

        // console.log(`MYX level: ${level}`);
        // Reset counts for all lower levels
        for (let i = level + 1; i < headingLevels.length; i++) {
          headingLevels[i] = 0;
        }

        // Construct the heading number (e.g., "8.3.2")
        const headingNumber = headingLevels.slice(0, level + 1).join(".");

        // console.log(`MYX headingNumber: ${headingNumber}`);

        let liText = `${headingNumber} ${paragraph.text}`;
        // console.log(`MYX liText: ${liText}`);




        const listItem = document.createElement("li");
        listItem.textContent = liText;
        listItem.style.cursor = "pointer";

        indentBasedOnHeading(paragraph, listItem);
        colorParagraphBasedOnSign(paragraph, listItem);

        listItem.addEventListener("click", (event) => navigateToHeading(context, index, event));

        listItem.addEventListener('contextmenu', async (event) => {
          event.preventDefault(); // Prevent the default context menu from appearing

          // Call your custom function here
          await customRightClickFunction(event, index);

          // You can use event.pageX and event.pageY to get the mouse position
          // console.log(`Right-clicked at position (${event.pageX}, ${event.pageY})`);

          return false; // Some browsers may require this to prevent the default action
        }, false);

        async function customRightClickFunction(event, paragraphIndex) {
          // Define your custom behavior for a right-click here
          console.log("Custom right-click function called.");

          event.preventDefault(); // Prevent the default context menu
          modifyParagraphSign(context, index, event);



          // const paragraphs = context.document.body.paragraphs;
          // paragraphs.load('text');
          // await context.sync();

          // const targetParagraph = paragraphs.items[paragraphIndex];
          // // addSignToHeading(targetParagraph); // Add a checkmark to the heading
          // targetParagraph.insertText("✓", Word.InsertLocation.end);
          // await context.sync();
          createFakeNavPane();

          // For example, show a custom context menu or perform other actions
        }

        headingsList.appendChild(listItem);
      }
    });
  }).catch((error) => console.error(error));
}

function modifyParagraphSign(context, paragraphIndex, event) {
  Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    await context.sync();

    let paragraph = paragraphs.items[paragraphIndex];
    let text = paragraph.text.trim();
    const symbols = ["✓", "✗", "⚠"];

    // Find existing symbols and their positions
    let foundSymbols = symbols.map(symbol => text.lastIndexOf(symbol));
    let maxPosition = Math.max(...foundSymbols);

    // Decide the action based on the found symbol
    if (maxPosition === -1) {
      // No symbol found, add the first symbol at the end with a space
      paragraph.insertText(" ✓", Word.InsertLocation.end);
    } else {
      // Remove duplicates and ensure only the last symbol is considered
      text = text.substring(0, maxPosition).replace(/\s+/g, '') + text.substring(maxPosition).replace(new RegExp(`[${symbols.join("")} ]`, "g"), "");
      const nextSymbolIndex = (foundSymbols.indexOf(maxPosition) + 1) % symbols.length;

      // If it's the last symbol, just remove it; otherwise, replace it with the next one
      if (nextSymbolIndex === 0) {
        // Just remove the last symbol (and the space before it if exists)
        paragraph.insertText(text, Word.InsertLocation.replace);
      } else {
        // Add the next symbol with a space
        text += ` ${symbols[nextSymbolIndex]}`;
        paragraph.insertText(text, Word.InsertLocation.replace);
      }
    }
    await context.sync();
    createFakeNavPane(); // Refresh the nav pane to reflect the updates
  });
}

function indentBasedOnHeading(paragraph, listItem) {
  const heading = paragraph.style;
  const level = parseInt(heading.split(" ")[1]);
  listItem.style.textIndent = level * 20 + "px";
  return level;
}

function colorParagraphBasedOnSign(paragraph, listItem) {
  if (paragraph.text.includes("✓")) {
    listItem.style.backgroundColor = "Green";
  }
  if (paragraph.text.includes("✗")) {
    listItem.style.backgroundColor = "Red";
  }
  if (paragraph.text.includes("⚠")) {
    listItem.style.backgroundColor = "Yellow";
  }
}
