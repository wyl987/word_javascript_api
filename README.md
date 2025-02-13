# Word JavaScript API Project

## Overview

This project utilizes the Microsoft Word JavaScript API to interact with Word documents. The add-in allows for detecting text formatting properties such as bold, underline, and font size of specific words within a document. It also includes automated tests written with Vitest to ensure proper functionality. The project provides an interactive user interface (UI) for displaying the detected text properties.

## Features

- Detects and checks text formatting (bold, underline, font size) for the first few words in a document.
- Handles empty or blank documents and displays appropriate messages.
- Provides real-time feedback on formatting through a user-friendly interface.
- Automated testing using Vitest to ensure correct behavior.

## Prerequisites

Before running the project, ensure that you have the following installed:

- **Node.js** (version 16 or above)
- **npm** or **yarn**
- **Microsoft Office Add-ins** enabled for Word

## Installation

Follow these steps to get the project running on your local machine:

1. **Clone the repository**:
    ```bash
    git clone https://github.com/wyl987/word_javascript_api.git
    ```

2. **Navigate to the project directory**:
    ```bash
    cd word_js_api_test
    ```

3. **Install dependencies**:
    ```bash
    npm install
    ```

## Running the Project

    ```bash
    npm start
    ```

## Testing

### Installation for Testing

To run tests, you'll need Vitest and testing libraries:

1. Install Vitest and other testing dependencies:
    ```bash
    npm install --save-dev vitest
    ```

### Running Tests

Run tests using the following command:

```bash
npm run test

### License
This project is licensed under the MIT License. See the LICENSE file for more details.
