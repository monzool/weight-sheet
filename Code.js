
const weight_cell = "D4"
const circumference_cell = "D6"
const date_cell = "D8"

// Function to Clear the User Form
function clearForm() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet
    var userForm = spreadsheet.getSheetByName("User form"); //declare a variable and set with the User Form worksheet

    // To create the instance of the user-interface environment to use the alert features
    var ui = SpreadsheetApp.getUi();

    // Display a dialog box with a title, message, and "Yes" and "No" buttons. The user can also
    // close the dialog by clicking the close button in its title bar.
    var response = ui.alert("Bekræft", 'Vil du slette alle data?',ui.ButtonSet.YES_NO);

    // Checking the user response and proceed with clearing the form if user selects Yes
    if (response == ui.Button.YES) {

        userForm.getRange(weight_cell).clear(); // Weight
        userForm.getRange(circumference_cell).clear(); // Circumference
        userForm.getRange(date_cell).clear(); // Date

        // Assigning white as default background color
        userForm.getRange(weight_cell).setBackground('#abebc6');
        userForm.getRange(circumference_cell).setBackground('#d5f5e3');
        userForm.getRange(date_cell).setBackground('#eafaf1');

        return true;
    }
}


//Declare a function to validate the entry made by user in UserForm
function validateEntry() {
    var spreadsheet= SpreadsheetApp.getActiveSpreadsheet();
    var userForm = spreadsheet.getSheetByName("User form");

    // To create the instance of the user-interface environment to use the messagebox features
    var ui = SpreadsheetApp.getUi();

    // Assigning white as default background color
    userForm.getRange(weight_cell).setBackground('#abebc6');
    userForm.getRange(circumference_cell).setBackground('#d5f5e3');
    userForm.getRange(date_cell).setBackground('#eafaf1');

    // Validating Employee ID
    if (userForm.getRange(weight_cell).isBlank() === true) {
        ui.alert("Ugyldig eller manglende vægt. Angiv i kilo. Eks: 70.5");
        userForm.getRange(weight_cell).activate();
        userForm.getRange(weight_cell).setBackground('#c39bd3');
        return false;
    }

    // Validating Employee Name
    else if(userForm.getRange(circumference_cell).isBlank() === true) {
        ui.alert("Ugyldig eller manglende omkreds. Angiv omkreds i centimeter. Eks:  100.5");
        userForm.getRange(circumference_cell).activate();
        userForm.getRange(circumference_cell).setBackground('#c39bd3');
        return false;
    }

    return true;
}


// Function to submit the data to user data sheet
function submitData() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var userForm = spreadsheet.getSheetByName("User form");
    var datasheet = spreadsheet.getSheetByName("User data");

    // To create the instance of the user-interface environment to use the messagebox features
    var ui = SpreadsheetApp.getUi();

    // Display a dialog box with a title, message, and "Yes" and "No" buttons. The user can also
    // close the dialog by clicking the close button in its title bar.
    var response = ui.alert("Gem", 'Vil du gemme disse data?',ui.ButtonSet.YES_NO);

    // Checking the user response and proceed with clearing the form if user selects Yes
    if (response == ui.Button.NO) {
        return;
    }

    // Validating the entry. If validation is true then proceed with transferring the data to Database sheet
    if (validateEntry() === true) {
        var blankRow = datasheet.getLastRow()+1; // Identify the next blank row

        datasheet.getRange(blankRow, 1).setValue(userForm.getRange(weight_cell).getValue()); // Weight
        datasheet.getRange(blankRow, 2).setValue(userForm.getRange(circumference_cell).getValue()); // Circumference
        datasheet.getRange(blankRow, 3).setValue(userForm.getRange(date_cell).getValue()); // Date

        // date function to update the current date and time as submittted on
        datasheet.getRange(blankRow, 7).setValue(new Date()).setNumberFormat('yyyy-mm-dd'); // Submitted On

        //get the email address of the person running the script and update as Submitted By
        datasheet.getRange(blankRow, 8).setValue(Session.getActiveUser().getEmail()); //Submitted By

        ui.alert(' "Registering gemt: ' + userForm.getRange(weight_cell).getValue() + ' [kg], '
                                         + userForm.getRange(circumference_cell).getValue() + ' [cm]');

        // Clear the data from the Data Entry Form
        userForm.getRange(weight_cell).clear();
        userForm.getRange(circumference_cell).clear();
        userForm.getRange(date_cell).clear();
    }
}


function reloadDatabase() {
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let user_data = spreadsheet.getSheetByName("User data");
    let mobile_data = spreadsheet.getSheetByName("Mobile data");
    let database = spreadsheet.getSheetByName("Database");

    database.clearContents();

    function toIsoDate(value) {
        const offset = value.getTimezoneOffset();
        value = new Date(value.getTime() - (offset*60*1000));
        value = value.toISOString().split('T')[0];
        return value
    }

    function collect_mobile_form_data(rows) {
      entries = []
      rows.slice(1).forEach(function(value) {
          console.log('Val:' + value);

          // Google forms deliver date as '09/01/2023 08.49.11'. Convert to iso8610 date
          entryDate = toIsoDate(value[0]);

          const entry = {
              date: entryDate,
              person: value[1],
              weight: value[2],
              circumference: value[3],
          }
          entries.push([entry.date, entry.person, entry.weight, entry.circumference])
      });

      return entries
    }

    function collect_user_form_data(rows) {
      entries = []
      rows.slice(1).forEach(function(value) {
          entryDate = value[2];
          if (entryDate === '') {
              entryDate = value[6]
          }
          entryDate = toIsoDate(entryDate);

          const entry = {
              weight: value[0],
              circumference: value[1],
              date: entryDate,
              person: value[7],
          }
          entries.push([entry.date, entry.person, entry.weight, entry.circumference])
      });

      return entries
    }

    let mobile_rows = mobile_data.getDataRange().getValues();
    mobile_entries = collect_mobile_form_data(mobile_rows)

    let user_rows = user_data.getDataRange().getValues();
    user_entries = collect_user_form_data(user_rows)

    entries = mobile_entries.concat(user_entries)

    // Sorting lexicographical. This works perfectly as date is first element, name next
    entries.sort()

    entries.forEach(function(entry) {
        database.appendRow(entry)
    });
}