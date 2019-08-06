# How to Automate Email Marketing using Google AppScripts?
This tutorial covers how to customized Google Slides for a bulk of customers in your Google Spreadsheets, and do a mass mailing with the slides as an attachment.

Here's an example of spreadsheets with customers info we are going to merge.
<img src="https://github.com/vanessaaleung/GoogleScriptsAutomation/blob/master/sample-screenshots/spreadsheet-data.png" alt="Sample Spreadsheet" width="500"/>

In the example, we have four columns: Customer to store customer's name, Email to save customer's email addresses, Message to store email contents for that customer, Status to save the sending status(SENT or not) of the email, to prevent sending duplicate messages.

## Part I: Customized Slides for All Customers
### Step 1: Setup
1. Create Tags in Your Slide

Open your slide template. For the element you want to insert texts, replace the content with a tag like `{{customer-name}}`. Make sure to use strings that are unlikely to occur normally.

<img src="https://github.com/vanessaaleung/GoogleScriptsAutomation/blob/master/sample-screenshots/tag-example.png" alt="Create Tags in Your Slide" width="500"/>


2. Enable Google APIs

In your slide, go to Tools -> Script editor -> Resources -> Advanced Google services -> Turn `Google Sheets API`, `Google Slides API`, `Drive API` on

<img src="https://github.com/vanessaaleung/GoogleScriptsAutomation/blob/master/sample-screenshots/api.png" alt="Enable Google APIs" width="600"/>

3. In Script editor, paste the code inside [copySlides.gs](https://github.com/vanessaaleung/GoogleScriptsAutomation/blob/master/copySlides.gs)


### Step 2: Replace with your codes
1. Replace Your Spreadsheet Id & Slide Id
You can easily get a Google document Id from its URL. For example, the Id of a document with URL `https://docs.google.com/presentation/d/12345/edit#slide=id.p` is `12345`.

2. Set the spreadsheet column you store the replace text to `customerName`.

3. Set the title of copied slides to `copyTitle`. In this example, it's set as `customerName presentation`.

4. Replace `{{YOUR_TAG}}` with the tag you created in Step 1.

### Step 3: Run Scripts
1. Click Save & Run 
2. Review Authorization Permissions
3. Now you should see slides copies & their pdf version in your Google Drive

## Part II: Customized Email for All Customers
### Step 1: Script Editor
In your spreadsheet, go to `Tools` -> `Script editor`, paste the codes in [sendEmails.gs](https://github.com/vanessaaleung/GoogleScriptsAutomation/blob/master/sendEmails.gs)

<img src="https://github.com/vanessaaleung/GoogleScriptsAutomation/blob/master/sample-screenshots/script-editor.png" alt="Script Editor" width="300"/>

### Step 2: Replace with your codes

Fetch dataReplace `startRow` with the index of the starting row you want to process, replace `numRows` with the number of rows you want to process.

2. Set `getRange(row, column, numRows,numColumns)` to fetch your data range. In this example, we fetch cells `A2:B4`.

3. Set `subject` to your expected email subject.

### Step 3: Run Scripts
1. Click Save & Run 
2. Review Authorization Permissions
3. Now emails with customized message & deck attachment should be sent to your customers.

<img src="https://github.com/vanessaaleung/GoogleScriptsAutomation/blob/master/sample-screenshots/email-sample.png" alt="Email Sample" width="500"/>
