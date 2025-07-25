# Auto-Complete Add-in for Microsoft Word

## Overview
This project is a Microsoft Word Add-in that provides real-time word suggestions as you type. The add-in dynamically tracks the words typed in the document and displays a dropdown of suggestions based on the current word being typed. It also includes a "Run" button to manually trigger the suggestion functionality.

## Features
- **Real-Time Suggestions**: Displays word suggestions dynamically while typing.
- **Dropdown at Cursor Position**: Suggestions appear near the cursor for seamless interaction.
- **Manual Trigger**: Use the "Run" button to manually fetch suggestions.
- **Top Suggestions**: Shows the top 5 suggestions based on the prefix of the currently typed word.

## How It Works
1. **Tracking Words**:
   - The add-in tracks all words typed in the document using the Word JavaScript API.
   - Words are stored in a global `Set` to avoid duplicates.

2. **Fetching Suggestions**:
   - Suggestions are fetched based on the prefix of the currently typed word.
   - The prefix is extracted dynamically from the text at the cursor position.

3. **Displaying Suggestions**:
   - A dropdown is displayed near the cursor using the `boundingRect` property of the selection.
   - Users can click on a suggestion to insert it into the document.

4. **Event Listener**:
   - A `keyup` event listener dynamically triggers the suggestion functionality while typing.

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/auto-complete-v2.git
```

## How to run this project

### Prerequisites

- Node.js (the latest LTS version). Visit the [Node.js site](https://nodejs.org/) to download and install the right version for your operating system. To verify that you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.
- Office connected to a Microsoft 365 subscription. You might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program), see [FAQ](https://learn.microsoft.com/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-) for details. Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try?rtc=1) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products).

### Run the add-in using Office Add-ins Development Kit extension

1. **Open the Office Add-ins Development Kit**
    
    In the **Activity Bar**, select the **Office Add-ins Development Kit** icon to open the extension.

1. **Preview Your Office Add-in (F5)**

    Select **Preview Your Office Add-in(F5)** to launch the add-in and debug the code. In the Quick Pick menu, select the option **Word Desktop (Edge Chromium)**.

    The extension then checks that the prerequisites are met before debugging starts. Check the terminal for detailed information if there are issues with your environment. After this process, the Word desktop application launches and sideloads the add-in.

1. **Stop Previewing Your Office Add-in**

    Once you are finished testing and debugging the add-in, select **Stop Previewing Your Office Add-in**. This closes the web server and removes the add-in from the registry and cache.

## Use the add-in project

The add-in project that you've created contains code for a basic task pane add-in.

## Explore the add-in code

To explore an Office add-in project, you can start with the key files listed below.

- The `./manifest.xml` file in the root directory of the project defines the settings and capabilities of the add-in.  <br>You can check whether your manifest file is valid by selecting **Validate Manifest File** option from the Office Add-ins Development Kit.
- The `./src/taskpane/taskpane.html` file contains the HTML markup for the task pane.
- The `./src/taskpane/taskpane.css` file contains the CSS that's applied to content in the task pane.
- The `./src/taskpane/taskpane.js` file contains the Office JavaScript API code that facilitates interaction between the task pane and the Word application.

## Troubleshooting

If you have problems running the add-in, take these steps.

- Close any open instances of Word.
- Close the previous web server started for the add-in with the **Stop Previewing Your Office Add-in** Office Add-ins Development Kit extension option.

If you still have problems, see [troubleshoot development errors](https://learn.microsoft.com//office/dev/add-ins/testing/troubleshoot-development-errors) or [create a GitHub issue](https://aka.ms/officedevkitnewissue) and we'll help you.  

For information on running the add-in on Word on the web, see [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

For information on debugging on older versions of Office, see [Debug add-ins using developer tools in Microsoft Edge Legacy](https://learn.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-legacy).

## Make code changes

All the information about Office Add-ins is found in our [official documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins). You can also explore more samples in the Office Add-ins Development Kit. Select **View Samples** to see more samples of real-world scenarios.

If you edit the manifest as part of your changes, use the **Validate Manifest File** option in the Office Add-ins Development Kit. This shows you errors in the manifest syntax.

## Engage with the team

Did you experience any problems? [Create an issue](https://aka.ms/officedevkitnewissue) and we'll help you out.

Want to learn more about new features and best practices for the Office platform? [Join the Microsoft Office Add-ins community call](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins-community-call).

## Copyright

Copyright (c) 2024 Microsoft Corporation. All rights reserved.

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**