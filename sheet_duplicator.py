/**
 * Opens the dialog with pre-filled values.
 * Sheet name = exact name of the last sheet (editable).
 */
function duplicateLastSheetWithInput() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var lastSheet = sheets[sheets.length - 1];
  var lastSheetName = lastSheet.getName();

  var props = PropertiesService.getDocumentProperties();
  var saved = props.getProperties();

  var prefill = {
    sheetName:   lastSheetName,
    challanNo:   saved.challanNo   || lastSheet.getRange('F59').getValue() || '',
    challanDate: saved.challanDate || formatDate(lastSheet.getRange('G59').getValue()),
    invoiceDate: saved.invoiceDate || formatDate(lastSheet.getRange('H59').getValue()),
    vehicleNo:   saved.vehicleNo   || lastSheet.getRange('I59').getValue() || '',
    house1st:    saved.house1st    || lastSheet.getRange('G62').getValue() || '',
    age1st:      saved.age1st      || lastSheet.getRange('G63').getValue() || '',
    birds1st:    saved.birds1st    || lastSheet.getRange('H64').getValue() || '',
    house2nd:    saved.house2nd    || lastSheet.getRange('J62').getValue() || '',
    age2nd:      saved.age2nd      || lastSheet.getRange('J63').getValue() || '',
    birds2nd:    saved.birds2nd    || lastSheet.getRange('J64').getValue() || ''
  };

  var html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial; padding: 20px; max-width: 420px;">
      <h3 style="margin-top:0;color:#1a73e8;">📋 Create New Sheet from Last Sheet</h3>
      <p style="color:#5f6368;font-size:13px;margin-top:0;">
        Source: <b>${lastSheetName}</b>
      </p>

      <label style="font-weight:600;font-size:13px;">Sheet Name * <span style="color:#1a73e8;font-weight:normal;">(edit if needed)</span></label>
      <input id="sheetName" type="text" value="${prefill.sheetName}"
             style="width:100%;padding:8px;margin-top:4px;font-size:14px;
                    border:2px solid #1a73e8;border-radius:6px;box-sizing:border-box;" />

      <hr style="margin:14px 0;border:none;border-top:1px solid #eee;" />

      <label style="font-weight:600;font-size:13px;">Challan No (F59)</label>
      <input id="challanNo" type="text" value="${prefill.challanNo}"
             style="width:100%;padding:8px;margin-top:4px;margin-bottom:8px;
                    border:1px solid #dadce0;border-radius:6px;box-sizing:border-box;" />

      <label style="font-weight:600;font-size:13px;">Challan Date (G59)</label>
      <input id="challanDate" type="date" value="${prefill.challanDate}"
             style="width:100%;padding:8px;margin-top:4px;margin-bottom:8px;
                    border:1px solid #dadce0;border-radius:6px;box-sizing:border-box;" />

      <label style="font-weight:600;font-size:13px;">Invoice Date (H59)</label>
      <input id="invoiceDate" type="date" value="${prefill.invoiceDate}"
             style="width:100%;padding:8px;margin-top:4px;margin-bottom:8px;
                    border:1px solid #dadce0;border-radius:6px;box-sizing:border-box;" />

      <label style="font-weight:600;font-size:13px;">Vehicle No (I59)</label>
      <input id="vehicleNo" type="text" value="${prefill.vehicleNo}"
             style="width:100%;padding:8px;margin-top:4px;margin-bottom:8px;
                    border:1px solid #dadce0;border-radius:6px;box-sizing:border-box;" />

      <hr style="margin:14px 0;border:none;border-top:1px solid #eee;" />

      <label style="font-weight:600;font-size:13px;">House #1st (G62)</label>
      <input id="house1st" type="text" value="${prefill.house1st}"
             style="width:100%;padding:8px;margin-top:4px;margin-bottom:8px;
                    border:1px solid #dadce0;border-radius:6px;box-sizing:border-box;" />

      <label style="font-weight:600;font-size:13px;">Age #1st (G63)</label>
      <input id="age1st" type="text" value="${prefill.age1st}"
             style="width:100%;padding:8px;margin-top:4px;margin-bottom:8px;
                    border:1px solid #dadce0;border-radius:6px;box-sizing:border-box;" />

      <label style="font-weight:600;font-size:13px;">Birds #1st (H64)</label>
      <input id="birds1st" type="text" value="${prefill.birds1st}"
             style="width:100%;padding:8px;margin-top:4px;margin-bottom:8px;
                    border:1px solid #dadce0;border-radius:6px;box-sizing:border-box;" />

      <label style="font-weight:600;font-size:13px;">House #2nd (J62)</label>
      <input id="house2nd" type="text" value="${prefill.house2nd}"
             style="width:100%;padding:8px;margin-top:4px;margin-bottom:8px;
                    border:1px solid #dadce0;border-radius:6px;box-sizing:border-box;" />

      <label style="font-weight:600;font-size:13px;">Age #2nd (J63)</label>
      <input id="age2nd" type="text" value="${prefill.age2nd}"
             style="width:100%;padding:8px;margin-top:4px;margin-bottom:8px;
                    border:1px solid #dadce0;border-radius:6px;box-sizing:border-box;" />

      <label style="font-weight:600;font-size:13px;">Birds #2nd (J64)</label>
      <input id="birds2nd" type="text" value="${prefill.birds2nd}"
             style="width:100%;padding:8px;margin-top:4px;margin-bottom:8px;
                    border:1px solid #dadce0;border-radius:6px;box-sizing:border-box;" />

      <div style="text-align:right;margin-top:16px;">
        <button onclick="google.script.host.close()"
                style="padding:8px 18px;background:#f1f3f4;color:#5f6368;
                       border:none;border-radius:6px;cursor:pointer;margin-right:6px;">
          Cancel
        </button>
        <button onclick="submitForm()"
                style="padding:8px 22px;background:#1a73e8;color:white;
                       border:none;border-radius:6px;cursor:pointer;font-weight:600;">
          Create Sheet
        </button>
      </div>

      <script>
        function submitForm() {
          var data = {
            sheetName:   document.getElementById('sheetName').value,
            challanNo:   document.getElementById('challanNo').value,
            challanDate: document.getElementById('challanDate').value,
            invoiceDate: document.getElementById('invoiceDate').value,
            vehicleNo:   document.getElementById('vehicleNo').value,
            house1st:    document.getElementById('house1st').value,
            age1st:      document.getElementById('age1st').value,
            birds1st:    document.getElementById('birds1st').value,
            house2nd:    document.getElementById('house2nd').value,
            age2nd:      document.getElementById('age2nd').value,
            birds2nd:    document.getElementById('birds2nd').value
          };
          google.script.run
            .withSuccessHandler(function(){ google.script.host.close(); })
            .withFailureHandler(function(err){ alert('Error: ' + err.message); })
            .processSheetName(data);
        }
      </script>
    </div>
  `)
  .setWidth(480)
  .setHeight(720);

  SpreadsheetApp.getUi().showModalDialog(html, 'Create New Sheet');
}

/**
 * Formats a date value into YYYY-MM-DD for HTML date input.
 */
function formatDate(value) {
  if (!value) return '';
  if (value instanceof Date) {
    var y = value.getFullYear();
    var m = ('0' + (value.getMonth() + 1)).slice(-2);
    var d = ('0' + value.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  return '';
}

/**
 * Creates the new sheet, writes values, and saves inputs for next time.
 */
function processSheetName(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var lastSheet = sheets[sheets.length - 1];
  var lastSheetName = lastSheet.getName();

  var newName = (data.sheetName || '').trim();

  if (newName === '') {
    SpreadsheetApp.getUi().alert('Error', 'Sheet name cannot be empty!',
                                 SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  if (newName !== lastSheetName && ss.getSheetByName(newName)) {
    SpreadsheetApp.getUi().alert('Error',
      'A sheet named "' + newName + '" already exists!',
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var newSheet = lastSheet.copyTo(ss);
  newSheet.setName(newName);

  newSheet.getRange('A1').setValue(newName);

  newSheet.getRange('F59').setValue(data.challanNo);
  if (data.challanDate)  newSheet.getRange('G59').setValue(new Date(data.challanDate));
  if (data.invoiceDate)  newSheet.getRange('H59').setValue(new Date(data.invoiceDate));
  newSheet.getRange('I59').setValue(data.vehicleNo);

  newSheet.getRange('G62').setValue(data.house1st);
  newSheet.getRange('G63').setValue(data.age1st);
  newSheet.getRange('H64').setValue(data.birds1st);

  newSheet.getRange('J62').setValue(data.house2nd);
  newSheet.getRange('J63').setValue(data.age2nd);
  newSheet.getRange('J64').setValue(data.birds2nd);

  PropertiesService.getDocumentProperties().setProperties({
    sheetName:   data.sheetName,
    challanNo:   data.challanNo,
    challanDate: data.challanDate,
    invoiceDate: data.invoiceDate,
    vehicleNo:   data.vehicleNo,
    house1st:    data.house1st,
    age1st:      data.age1st,
    birds1st:    data.birds1st,
    house2nd:    data.house2nd,
    age2nd:      data.age2nd,
    birds2nd:    data.birds2nd
  });

  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(sheets.length + 1);

  SpreadsheetApp.getUi().alert('Success',
    'Sheet "' + newName + '" created successfully!\n\n' +
    'The values you entered are now saved for next time.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Optional: clear saved pre-fill values.
 */
function clearSavedValues() {
  PropertiesService.getDocumentProperties().deleteAllProperties();
  SpreadsheetApp.getUi().alert('Cleared',
    'Saved pre-fill values have been cleared.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Menu
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Automation')
    .addItem('Duplicate Last Sheet', 'duplicateLastSheetWithInput')
    .addSeparator()
    .addItem('Clear Saved Pre-fill Values', 'clearSavedValues')
    .addToUi();
}
