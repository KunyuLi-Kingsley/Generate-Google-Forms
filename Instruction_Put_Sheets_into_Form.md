# Instruction: Put Google Sheets Data into Google Forms

## 1. Feature Overview

This tool automatically reads questionnaire data from Google Sheets and creates questions in Google Forms based on the spreadsheet content.

## 2. Spreadsheet Format

### 2.1 Required Google Sheets format

The script expects the Google Sheet to use a fixed four-column format:

| Section | Question | Type | Options |
|---|---|---|---|
| Cycling Usage & User Profile | What is your age? | Multiple choice | Under 18, 18-34, 35-54, 55+ |

Column descriptions:

- **Section**: The category or section of the questionnaire.
- **Question**: The question text.
- **Type**: The question type.
- **Options**: The answer choices. Separate multiple choices with English commas.

A suggested workflow is to copy the prompt below into a text-generation AI tool, change the topic and number of questions to fit your needs, generate a table, and then review or edit the results before pasting them into Google Sheets.

### Prompt template

```text
请帮我生成30个市场调研的题目，题型对应四大咨询公司的题型，涵盖单选、多选、一道开放题，主题是【跑步包袋市场调研】请严格按照以下格式输出表格内容:
Section | Question | Type | Options

Section: 请为每组题目提供合适的分类标题。
Type: 固定为 Short answer、Multiple choice 或 Checkboxes（三选一，依据题目内容自动选择）。
Options: 如果是选择题或多选题，请提供4个选项，用英文逗号分隔。
表格格式请使用表格，方便我直接复制进 Google Sheets，请开始生成题目。
```

### 2.2 Question column

Enter the question text in the second column. For example:

```text
What is your primary cycling purpose?
```

The script automatically adds question numbers in the generated Google Form. Do **not** manually add numbers such as `1.`, `2.`, or `3.` in the Sheet.

For example, the Sheet can contain:

```text
What is your primary cycling purpose?
How often do you ride a bicycle?
What type of bicycle do you mainly use?
```

The generated Google Form will display:

```text
1. What is your primary cycling purpose?
2. How often do you ride a bicycle?
3. What type of bicycle do you mainly use?
```

### 2.3 Type column

Use one of the following values in the `Type` column:

| Value in Google Sheets | Google Forms question type |
|---|---|
| `Multiple choice` | Multiple-choice question (single selection) |
| `Checkboxes` | Checkbox question (multiple selections) |
| `Short answer` | Short-answer text question |

The script converts the type text to lowercase before checking it, so capitalization does not matter for these supported values.

### 2.4 Options column

Enter multiple answer options in the fourth column, separated by English commas.

Example:

```text
Daily commuting, Fitness and exercise, Leisure and recreation, Long-distance cycling
```

The script splits the text at commas, trims extra spaces, removes empty entries, and removes duplicate options. Google Forms then displays the choices as separate options.

For a checkbox question, for example:

```text
Commuting, Fitness, Leisure, Long-distance cycling
```

the form displays four checkboxes:

- Commuting
- Fitness
- Leisure
- Long-distance cycling

For a `Short answer` question, the `Options` cell can be left blank.

Example:

| Section | Question | Type | Options |
|---|---|---|---|
| Cycling Usage | Please describe your cycling habits. | Short answer | |

## 3. How to Use the Script

Only three configuration values need to be set in the script:

- `SHEET_NAME`: The name of the individual worksheet tab inside the Google Sheet. The example in the original document uses `aa`.
- `FORM_ID`: The ID of the target Google Form. Create a blank form first, then copy the ID from the URL between `/d/` and `/edit`.
- `SHEET_ID`: The ID of the Google Spreadsheet. Copy the ID from its URL between `/d/` and `/edit`.

Example URLs:

```text
Google Form:
https://docs.google.com/forms/d/FORM_ID/edit

Google Sheet:
https://docs.google.com/spreadsheets/d/SHEET_ID/edit
```

Replace the three values in the script with the IDs and sheet name for your own files, then run the function.

### Script link

The original document provides this Google Apps Script project:

https://script.google.com/home/projects/1QJtQRBsRCNot4mbgP5rCkQZeY3cYIdtDO-Goslnwu4ro_ZZN26FPbY8s/edit

If the link is unavailable, use the code below.

## 4. Google Apps Script Code

> **Important:** The script below deletes all existing items in the target Google Form before creating new questions from the Sheet. Use a blank form or a form whose existing questions you are prepared to replace. Check the three configuration values before running it.

```javascript
function populateDropdownsFromSheet() {
  // Values to configure
  // =========================
  const SHEET_NAME = "REPLACE_WITH_SHEET_TAB_NAME";
  const FORM_ID = "REPLACE_WITH_GOOGLE_FORM_ID";
  const SHEET_ID = "REPLACE_WITH_GOOGLE_SHEET_ID";
  // =========================

  const form = FormApp.openById(FORM_ID);

  Logger.log(
    "Opened Form: " + form.getTitle()
  );

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error(
      "Sheet not found: " + SHEET_NAME
    );
  }

  let data = sheet
    .getDataRange()
    .getDisplayValues();

  // Remove the header row
  data.shift();

  const oldItems = form.getItems();

  Logger.log(
    "Deleting old items: " + oldItems.length
  );

  for (
    let i = oldItems.length - 1;
    i >= 0;
    i--
  ) {
    form.deleteItem(oldItems[i]);
  }

  // Wait for deletions to synchronize
  Utilities.sleep(5000);

  let currentSection = "";
  let questionCount = 0;

  data.forEach(function(row, index) {
    let section = cleanText(row[0]);
    let originalQuestion = cleanText(row[1]);
    let type = cleanText(row[2]).toLowerCase();
    let optionText = cleanText(row[3]);

    if (originalQuestion === "") {
      return;
    }

    if (section !== currentSection) {
      let page = form.addPageBreakItem();

      Utilities.sleep(1500);
      page.setTitle(section);
      Utilities.sleep(1500);

      currentSection = section;
    }

    // Add question numbering
    let question =
      (questionCount + 1) + ". " + originalQuestion;

    // Process options
    let options = [];

    if (optionText !== "") {
      options = optionText
        .split(",")
        .map(function(x) {
          return cleanText(x);
        })
        .filter(function(x) {
          return x !== "";
        });

      // Remove duplicate options
      options = [...new Set(options)];
    }

    // Create the Google Forms item
    let item;

    try {
      // Single-choice question
      if (type === "multiple choice") {
        item = form.addMultipleChoiceItem();

        Utilities.sleep(1500);
        item.setTitle(question);

        Utilities.sleep(1500);
        item.setChoiceValues(options);
        item.setRequired(true);
      }

      // Multiple-selection question
      else if (type === "checkboxes") {
        item = form.addCheckboxItem();

        Utilities.sleep(1500);
        item.setTitle(question);

        Utilities.sleep(1500);
        item.setChoiceValues(options);
        item.setRequired(true);
      }

      // Short-answer question
      else if (type === "short answer") {
        item = form.addTextItem();

        Utilities.sleep(1500);
        item.setTitle(question);
      }

      // Default to a text question for unsupported types
      else {
        item = form.addTextItem();

        Utilities.sleep(1500);
        item.setTitle(question);
      }

      // Final title verification
      Utilities.sleep(2000);

      if (item.getTitle() !== question) {
        Logger.log(
          "Title missing; writing it again: " + question
        );

        item.setTitle(question);
        Utilities.sleep(1500);
      }

      questionCount++;

      Logger.log(
        "Created successfully: " +
        questionCount +
        " | " +
        question
      );
    } catch (e) {
      Logger.log(
        "Failed to create: " + question
      );

      Logger.log(e.message);
    }
  });

  Logger.log("========================");
  Logger.log(
    "Total questions created: " + questionCount
  );
  Logger.log(
    "Form items: " + form.getItems().length
  );
}

// =========================
// Text cleaning
// =========================
function cleanText(text) {
  if (text === null || text === undefined) {
    return "";
  }

  return String(text)
    .replace(/[\n\r]/g, " ")
    .replace(/\s+/g, " ")
    .replace(/\u00A0/g, " ")
    .trim();
}
```

## 5. Run the Script

1. Open the Google Apps Script project.
2. Paste the code into the editor, or use the provided script project if you have access.
3. Update `SHEET_NAME`, `FORM_ID`, and `SHEET_ID`.
4. Make sure the Google Sheet has the four required columns and a header row.
5. Select `populateDropdownsFromSheet` in the function selector.
6. Click **Run**.
7. If prompted, review and grant the required Google permissions.
8. Check the execution log for the number of questions created and any errors.
9. Open the target Google Form and verify the sections, question titles, question types, and answer options.

## 6. Important Notes

- The first row of the spreadsheet is treated as the header and removed before processing.
- Rows with an empty `Question` cell are skipped.
- A new page-break section is created whenever the `Section` value changes from the previous processed row. Group related questions together in the Sheet.
- The script numbers questions automatically.
- `Multiple choice` and `Checkboxes` questions are marked as required. `Short answer` questions are not explicitly marked as required.
- Answer choices are separated by English commas. If an answer option itself contains a comma, this simple format will split it into separate options.
- Unknown or misspelled question types fall back to text questions.
- The script deletes existing form items before rebuilding the form. Use caution if the form already contains questions or other items.
- The script includes several `Utilities.sleep()` calls to allow operations to complete. Execution time may vary with the number of questions and Google Apps Script limits.
