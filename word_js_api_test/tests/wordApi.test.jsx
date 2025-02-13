import { describe, it, expect, vi, beforeEach, beforeAll } from "vitest";
// import { run } from "../src/taskpane/taskpane";

beforeAll(() => {
  globalThis.Word = {
    run: async (callback) => {
      // Provide a dummy context. This will be overridden in beforeEach.
      return await callback({
        document: {
          body: {
            getRange: vi.fn(),
            search: vi.fn(),
          },
        },
        sync: async () => {},
      });
    },
  };
  globalThis.Office = {
    onReady: vi.fn((callback) => {
      // Simulate Word environment being ready
      callback({ host: 'Word' });
    }),
    HostType: {
      Word: 'Word',
      Excel: 'Excel',
      PowerPoint: 'PowerPoint',
    },
  };
  // Create the required DOM elements
  document.body.innerHTML = `
    <div id="sideload-msg"></div>
    <div id="app-body"></div>
    <div id="header-title"></div>
    <button id="run"></button>
  `;
});

describe("Test run function", () => {
  let context;
  let body;
  let range;

  beforeEach(() => {
    // Create a new Word context mock for each test.
    context = {
      document: {
        body: {
          getRange: vi.fn(),
          search: vi.fn(),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    body = context.document.body;

    range = {
      text: '',
      load: vi.fn(),
    };
    body.getRange.mockReturnValue(range);

    // Override the global Word object with our test-specific version.
    globalThis.Word = {
      run: async (callback) => await callback(context),
    };

    // Reset the header-title element for a clean test slate.
    const headerTitle = document.getElementById("header-title");
    headerTitle.innerHTML = "";
  });

  it('should detect bold, underline, and font size correctly for 3 words', async () => {
    const { run } = await import("../src/taskpane/taskpane");
    
    range.text = "Good Afternoon mate";

    // Mocking Word range objects with font properties
    const firstWordRange = { font: { bold: true }, load: vi.fn(), };
    const secondWordRange = { font: { underline: 'Single' }, load: vi.fn(), };
    const thirdWordRange = { font: { size: 15 }, load: vi.fn(), };

    body.search.mockImplementation((searchText) => {
      if (searchText === 'Good') return { getFirstOrNullObject: vi.fn(() => firstWordRange) };
      if (searchText === 'Afternoon') return { getFirstOrNullObject: vi.fn(() => secondWordRange) };
      if (searchText === 'mate') return { getFirstOrNullObject: vi.fn(() => thirdWordRange) };
      return { getFirstOrNullObject: vi.fn(() => null) };
    });

    await run();

    // Expect the header text to contain the correct results
    const headerTitle = document.getElementById("header-title");
    expect(headerTitle.innerHTML).toContain("First word is bold: True");
    expect(headerTitle.innerHTML).toContain("Second word has underline: True");
    expect(headerTitle.innerHTML).toContain("Font size of the third word: 15");
  });

  it('should detect not bold, no underline, and font size correctly for 3 words', async () => {
    const { run } = await import("../src/taskpane/taskpane");
    
    range.text = "Good Afternoon mate";

    // Mocking Word range objects with font properties
    const firstWordRange = { font: { bold: false }, load: vi.fn(), };
    const secondWordRange = { font: { underline: 'None' }, load: vi.fn(), };
    const thirdWordRange = { font: { size: 15 }, load: vi.fn(), };

    body.search.mockImplementation((searchText) => {
      if (searchText === 'Good') return { getFirstOrNullObject: vi.fn(() => firstWordRange) };
      if (searchText === 'Afternoon') return { getFirstOrNullObject: vi.fn(() => secondWordRange) };
      if (searchText === 'mate') return { getFirstOrNullObject: vi.fn(() => thirdWordRange) };
      return { getFirstOrNullObject: vi.fn(() => null) };
    });

    await run();

    // Expect the header text to contain the correct results
    const headerTitle = document.getElementById("header-title");
    expect(headerTitle.innerHTML).toContain("First word is bold: False");
    expect(headerTitle.innerHTML).toContain("Second word has underline: False");
    expect(headerTitle.innerHTML).toContain("Font size of the third word: 15");
  });

  it('should correctly display "The document is blank." when the document is blank', async () => {
    const { run } = await import("../src/taskpane/taskpane");
    
    range.text = "";

    await run();

    // Expect the header text to contain the correct results
    const headerTitle = document.getElementById("header-title");
    expect(headerTitle.innerHTML).toContain("The document is blank.");
  });

  it('should handle one word with bold correctly', async () => {
    const { run } = await import("../src/taskpane/taskpane");
    
    range.text = "Good";

    // Mocking Word range objects with font properties
    const firstWordRange = { font: { bold: false }, load: vi.fn(), };

    body.search.mockImplementation((searchText) => {
      if (searchText === 'Good') return { getFirstOrNullObject: vi.fn(() => firstWordRange) };
      return { getFirstOrNullObject: vi.fn(() => null) };
    });

    await run();

    // Expect the header text to contain the correct results
    const headerTitle = document.getElementById("header-title");
    expect(headerTitle.innerHTML).toContain("First word is bold: False");
    expect(headerTitle.innerHTML).toContain("Second word has underline: False");
    expect(headerTitle.innerHTML).toContain("Font size of the third word: Not available");
  });

  it('should handle two words both bold and underlined correctly', async () => {
    const { run } = await import("../src/taskpane/taskpane");
    
    range.text = "Good Afternoon mate";
    // Mocking Word range objects with font properties
    const firstWordRange = { font: { bold: true }, load: vi.fn(), };
    const secondWordRange = { font: { underline: 'Single' }, load: vi.fn(), };

    body.search.mockImplementation((searchText) => {
      if (searchText === 'Good') return { getFirstOrNullObject: vi.fn(() => firstWordRange) };
      if (searchText === 'Afternoon') return { getFirstOrNullObject: vi.fn(() => secondWordRange) };
      return { getFirstOrNullObject: vi.fn(() => null) };
    });

    await run();

    // Expect the header text to contain the correct results
    const headerTitle = document.getElementById("header-title");
    expect(headerTitle.innerHTML).toContain("First word is bold: True");
    expect(headerTitle.innerHTML).toContain("Second word has underline: True");
    expect(headerTitle.innerHTML).toContain("Font size of the third word: Not available");
  });
  
});
  