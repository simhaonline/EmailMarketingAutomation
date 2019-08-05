# How to merge Google Spreadsheets Data to Google Slides & Email
## Part I: Merge Spreadsheets text to Slides
### Step 1: Setup
1. Create Tags in Your Slide

Open your slide template. For the element you want to insert texts, replace the content with a tag like `{{customer-name}}`. Make sure to use strings that are unlikely to occur normally.

2. Enable Google APIs

In your slide, go to Tools -> Script editor -> Resources -> Advanced Google services -> Turn `Google Sheets API`, `Google Slides API`, `Drive API` on
Turn Google Sheets API, Google Slides API, Drive API on

3. In Script editor, paste the code inside `copySlides.gs`


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

## Part II: Merge Spreadsheets to Email
### Step 1: Get to Know Your Spreadsheet Data
In the example, we have four columns: Customer to store customer's name, Email to save customer's email addresses, Message to store email contents for that customer, Status to save the sending status(SENT or not) of the email, to prevent sending duplicate messages.
In this example, we have four columns

### Step 2: Script Editor
In your spreadsheet, go to `Tools` -> `Script editor`, paste the codes in `sendEmails.gs`

### Step 3: Replace with your codes

Fetch dataReplace `startRow` with the index of the starting row you want to process, replace `numRows` with the number of rows you want to process.

2. Set `getRange(row, column, numRows,numColumns)` to fetch your data range. In this example, we fetch cells `A2:B4`.

Set `subject` to your expected email subject.
Now emails with customized message & deck attachment should be sent to your customers. Here's an example:
Example email

### Step 3: Run Scripts
1. Click Save & Run 
2. Review Authorization Permissions
3. Now emails with customized message & deck attachment should be sent to your customers.
