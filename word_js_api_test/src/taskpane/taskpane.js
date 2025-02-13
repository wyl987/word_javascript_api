/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // grab the current doc and load the first few words from the body
    const body = context.document.body;
    const range = body.getRange();
    range.load("text");
    await context.sync();

    if (range.text.trim() === "") {
      const headerTitle = document.getElementById("header-title");
      headerTitle.textContent = "The document is blank.";
      return;
    }

    // Split the document text into words and grab the first 3 words
    const words = range.text.split(/\s+/).map(word => word.replace(/[^\w\s]/g, '')).filter(word => word.length > 0).slice(0, 3);

    if (words.length === 0) {
      const headerTitle = document.getElementById("header-title");
      headerTitle.textContent = "The document doesn't have any words.";
      return;
    }

    let firstWordRange = null,
      secondWordRange = null,
      thirdWordRange = null;
    if (words.length >= 1) firstWordRange = body.search(words[0], { matchCase: true }).getFirstOrNullObject();
    if (words.length >= 2) secondWordRange = body.search(words[1], { matchCase: true }).getFirstOrNullObject();
    if (words.length >= 3) thirdWordRange = body.search(words[2], { matchCase: true }).getFirstOrNullObject();

    let isBold = false;
    let hasUnderline = false;
    let fontSize = "Not available";

      // Check if word ranges exist and load font properties
    if (firstWordRange) {
      firstWordRange.load("font/bold");
      await context.sync();
      isBold = firstWordRange.font.bold;
    }

    if (secondWordRange) {
      secondWordRange.load("font/underline");
      await context.sync();
      hasUnderline = secondWordRange.font.underline !== "None" ? true : false; 
    }

    if (thirdWordRange) {
      thirdWordRange.load("font/size");
      await context.sync();
      fontSize = thirdWordRange.font.size || "Not available";
    }

    // Display results in the UI
    const headerTitle = document.getElementById("header-title");
    headerTitle.textContent = "Results:";

    headerTitle.innerHTML = `
      First word is bold: ${isBold ? "True" : "False"}<br>
      Second word has underline: ${hasUnderline ? "True" : "False"}<br>
      Font size of the third word: ${fontSize}
    `;
  });
}
