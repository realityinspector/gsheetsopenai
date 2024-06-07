# gsheetsopenai
google sheets Apps Script for openai 

MIT LICENSE 

Google Sheets OpenAI Integration

This guide will help you integrate the OpenAI API into Google Sheets using Google Apps Script. The script provided will allow you to generate text using the OpenAI API based on the contents of the row of the active cell.
Features

    Combines content from the active row (excluding the active cell) into a single prompt.
    Sends the prompt to the OpenAI API and receives a response.
    Inserts the generated text into the active cell.

Setup Instructions
1. Open Google Sheets

Open the Google Sheets document where you want to integrate the OpenAI API.
2. Open Apps Script

    Click on Extensions in the top menu.
    Select Apps Script.

3. Delete Default Code

If there is any default code in the script editor, delete it.
4. Paste the Script

Copy and paste the following script into the Apps Script editor:

5. Set Your OpenAI API Key

Replace 'sk-' in the apiKey variable with your actual OpenAI API key.
6. Save the Script

    Click the floppy disk icon or press Ctrl + S to save the script.
    Name the project if prompted.

7. Run the Script

    Go back to your Google Sheets.
    You will now see a new menu item called AI.
    Select Generate Text from the AI menu to run the script.

Usage

    Select the cell where you want the generated text to appear.
    Ensure that there is content in the other cells of the same row.
    Click on AI > Generate Text.

The script will collect all content from the active row (excluding the active cell), create a prompt, send it to the OpenAI API, and display the generated response in the active cell.
Customization

You can customize the following parameters in the script:

    temperature: Adjusts the randomness of the generated text.
    max_tokens: Sets the maximum number of tokens in the generated text.
    top_p: Adjusts the diversity of the generated text.

Adjust these values according to your needs for different text generation behaviors.
