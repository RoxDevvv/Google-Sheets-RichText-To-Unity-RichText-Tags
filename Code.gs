function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Unity Formater Menu')
    .addItem('Toggle Formatted Text', 'toggleFormattedText')
    .addToUi();
}

function toggleFormattedText() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getActiveRange();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();


  for (var i = 0; i < numRows; i++) {
    for (var j = 0; j < numCols; j++) {
      var Cell = range.getCell(i + 1, j + 1);

      var cellValue = Cell.getValue(); // Get the content of the cell
      var isFormatted = /<[a-z][\s\S]*>/i.test(cellValue);
      var newText = isFormatted ? stripHtmlTags(cellValue) : formattedTextFromRichText(Cell);
      Cell.setValue(newText);
      Logger.log(
        isFormatted
          ? 'Removed formatting from cell '
          : 'Applied formatting to cell '
      );
    }
  }
}

function stripHtmlTags(text) {
  // Use a regular expression to remove HTML tags
  return text.replace(/<[^>]*>/g, '');
}

function formattedTextFromRichText(srange) {
  var styleRuns = srange.getRichTextValue().getRuns();
  var styledText = '';

  styleRuns.forEach(function (run) {
    var text = run.getText();
    var style = run.getTextStyle();
    var linkUrl = run.getLinkUrl();

    // Apply all formatting using Unity Rich Text tags
    if (style.isStrikethrough()) text = '<s>' + text + '</s>';
    if (style.isBold()) text = '<b>' + text + '</b>';
    if (style.isItalic()) text = '<i>' + text + '</i>';

    var color = style.getForegroundColor();
    var size = style.getFontSize();

    if (color && color !== '#000000' && (!linkUrl || color !== '#1155cc')) {
      // Use Unity Rich Text color tag
      text = '<color=' + color + '>' + text + '</color>';
    }

    if (size && size > 10) {
      // Use Unity Rich Text size tag
      text = '<size=' + size + '>' + text + '</size>';
    }

    if (style.isUnderline() && !linkUrl) text = '<u>' + text + '</u>';

    if (linkUrl) {
      // Unity Rich Text doesn't support hyperlinks directly, so you may need to handle them differently
      // You can create your own custom formatting for hyperlinks here
      // For example, you can use a custom tag like <link=URL>Text</link>
      text = '<link=' + linkUrl + '>' + text + '</link>';
    }

    styledText += text;
  });

  // Unity Rich Text doesn't support lists or certain advanced formatting directly,
  // so you may need to customize the output further based on your needs
  return styledText;
}
