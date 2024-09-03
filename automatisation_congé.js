function main(workbook: ExcelScript.Workbook) {
    
  // Getting the Follow Up File sheet
  let follow_up_file_sheet = workbook.getWorksheet("Follow Up File");
  let follow_up_file_date_values_range = follow_up_file_sheet.getRange("K5:KK5");

  // Get congé cell 
  let full_day_off_cell = follow_up_file_sheet.getRange("G2");
  let half_day_off_cell = follow_up_file_sheet.getRange("E2");
  let half_day_val = 0.5

  // Getting temp sheet
  let temp_sheet = workbook.getWorksheet("Temp");
  let status_values = temp_sheet.getRange("P1:P1000").getValues();
  let start_date_values = temp_sheet.getRange("K1:K1000").getValues();
  let end_date_values = temp_sheet.getRange("M1:M1000").getValues();
  let nbr_days_values = temp_sheet.getRange("Q1:Q1000").getValues();

    // // Get the ranges for J24 and M24.
    // let cellJ24 = selectedSheet.getRange("J24");
    // let cellM24 = selectedSheet.getRange("M24");

    // // Copy value and format from J24 to M24.
    // let cellJ24Value = cellJ24.getValue();
    // cellM24.copyFrom(cellJ24);

  // Finding the number of rows in Temp
  let tota_rows = getTotalRowsCount(temp_sheet);
  
  let start_index = 1;  // start_index =1 to skip column names row
  // highlightRow(temp_sheet, 1);

  // getEmployeeRowIndex(follow_up_file_date_values_range, )

  // Creating for loop to iterate through each row
  for (let i=start_index; i<= tota_rows-1; i++){

    let half_day_exists = false
    // Check if for the current row, the column P (status) has the value "Cancelled". If tha's the case continue. 
    let current_status = status_values[i][0];

    if (current_status == "Cancelled"){
      highlightRow(temp_sheet, i);
      console.log("Skipping row");
      continue;
    }
    
    // Otherwise :
    // Get start date (column K) and convert to save format as the date of Follow Up File (remove year)
    let current_start_date_serial = Number(start_date_values[i][0]);
    let current_start_date = convertExcelDateToJSDate(current_start_date_serial);
    let current_formatted_start_date = formatDateToDDMM(current_start_date);

    // TODO: Get end date (column M) and convert to save format as the date of Follow Up File (remove year)
    let current_end_date_serial = Number(end_date_values[i][0]);
    let current_end_date = convertExcelDateToJSDate(current_end_date_serial);
    let current_formatted_end_date = formatDateToDDMM(current_end_date);

    // TODO: Get number of days (column Q). If decimal (% =0.5), set half_day_exit = true
    let current_number_of_days = nbr_days_values[i][0];
    if (hasDecimalOfPointFive(current_number_of_days)) {
      half_day_exists = true;
      console.log("Half day exists: " + current_number_of_days)
    }

    // TODO: Find the start date column index (or string) in Follow Up File sheet
    let current_start_date_index = getIndex(follow_up_file_sheet, follow_up_file_date_values_range, current_formatted_start_date);
    
    // TODO: Find the end date column index (or string) in Follow Up File sheet
    let current_end_date_index = getIndex(follow_up_file_sheet, follow_up_file_date_values_range, current_formatted_end_date);
    // TODO: If start and end date not found, raise, highlit current row in red and skip iteration
    if (current_start_date_index == -1 || current_end_date_index == -1){
      highlightRow(temp_sheet, i);
      console.log("Skipping row");
      continue;
    }

    // TODO: Find row of current employee in Follow Up File sheet
    
    // TODO: Copy C from start to end date (while loop )

    // TODO: if half_day_exists, copy HD to last cell in range

  }
}


/**
 * Convert an Excel date serial number to a JavaScript Date object.
 * @param {number} serial - The Excel date serial number.
 * @returns {Date} - The JavaScript Date object.
 */
function convertExcelDateToJSDate(serial: number): Date {
  const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // Excel's base date (December 30, 1899)
  const jsDate = new Date(excelEpoch.getTime() + serial * 86400000); // 86400000 ms in a day
  return jsDate;
}

/**
 * Format a JavaScript Date object to a string in dd/MM format.
 * @param {Date} date - The JavaScript Date object.
 * @returns {string} - The formatted date string in dd/MM format.
 */
function formatDateToDDMM(date: Date): string {
  let day = date.getUTCDate().toString().padStart(1, '0');
  let month = (date.getUTCMonth() + 1).toString().padStart(1, '0'); // getUTCMonth() returns month from 0-11
  return `${day}/${month}`;
}

/**
 * Function to get total number of rows in the given sheet
 * @param {ExcelScript.Worksheet} sheet - The sheet to count rows in.
 * @returns {number} - Total number of rows with data.
 */
function getTotalRowsCount(sheet: ExcelScript.Worksheet): number {
  let usedRange = sheet.getUsedRange();
  return usedRange.getRowCount();
}

function getIndex(sheet: ExcelScript.Worksheet, range: ExcelScript.Range, target_value: string): number {
  let found_index = -1;  // Initialize with -1 to indicate not found
  let range_values = range.getValues()[0]

  for (let i = 0; i < range_values.length; i++) {
    if (range_values[i] === target_value) {
      found_index = i;
      // console.log("Index found");
      break;
    }
  }

  // // Log the result
  // if (found_index !== -1) {
  //   console.log(`The value "${target_value}" is found at index ${found_index} in the range Z5:ZZ5.`);
  // } else {
  //   console.log(`The value "${target_value}" is not found in the range Z5:ZZ5.`);
  // }
  // if (found_index == -1) {  // Verifying if date or name was found
  //   // If the flag is still false, none of the conditions were true
  //   throw new Error(`Date not found : ${target_value}`);
  // }
  return found_index;
}

function hasDecimalOfPointFive(value: number): Boolean {
  // Check if the fractional part of the number is 0.5
  if (Math.abs(value - Math.floor(value)) === 0.5) {
    return true;
  } else {
    return false;
  }
}

function highlightRow(sheet: ExcelScript.Worksheet, rowIndex: number) {
  // Define the range for the entire row
  let used_range = sheet.getUsedRange();
  let highlight_range = sheet.getRangeByIndexes(rowIndex, 0, 1, used_range.getColumnCount());
  // Get the RangeFill object.
  let fill = highlight_range.getFormat().getFill();
  // Set the fill color to yellow.
  fill.setColor("FFFF00");
  console.log("Row " + rowIndex + " highlighted.");
}

// Function that returns the row index of an employee name
function getEmployeeRowIndex(worksheet: ExcelScript.Worksheet, lookup_range: ExcelScript.Range, desired_name: string): number {
  // Getting the values from the lookup range
  let lookup_values = lookup_range.getValues();
  // Getting the index of the lookup value
  let row_index = lookup_values.findIndex(row => row[0] === desired_name);
  // Checking if the desired value was found
  if (row_index === -1) {
    throw new Error(`Employee name "${desired_name}" not found in the specified range.`);
  }
  return row_index;
}
