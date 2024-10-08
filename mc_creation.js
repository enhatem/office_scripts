function main(workbook: ExcelScript.Workbook) {
  // Getting the worksheets that will be used to duplicate the script 
  let config_sheet = workbook.getWorksheet("config");
  let mc_template_sheet = workbook.getWorksheet("MC_Template")
  // Getting the Start and End Date of the current period
  let start_date = config_sheet.getRange("C13").getValue();
  let end_date = config_sheet.getRange("C14").getValue();
  // Getting the relevant relative start and end row indices
  let relative_start_and_end_row_indices = getRelevantRelativeStartandEnd(config_sheet, start_date, end_date);
  // Getting the BLs list that will be used during the iteration process
  let bl_tabs_list = getListByName(config_sheet, relative_start_and_end_row_indices, "Onglet BL");
  // Getting the dates of the periods
  let dates_list = getListByName(config_sheet, relative_start_and_end_row_indices, "Dates des periodes");
  // Creating total database for workpackages
  let databases = createDatabase(workbook, bl_tabs_list);
  // Extracting total and individual databases objects
  let total_db = databases.totalDatabases;
  let individual_db = databases.individualDatabases;
  let desired_mc_worksheet_name = getDesiredMcWorksheetName(config_sheet, start_date);
  let mc_worksheet = workbook.getWorksheet(desired_mc_worksheet_name);
  // Checking if mc_worksheet exists
  if (mc_worksheet){  // If a tab with the desired MC worksheet name exists, update worksheet
    console.log("The tab " + desired_mc_worksheet_name + " exists. Updating existing sheet...");
    updateMcWorksheet(mc_worksheet, total_db, individual_db, dates_list);
  }else{  // Otherwise, create a new MC worksheet
    console.log("The tab " + desired_mc_worksheet_name + " does not exist, creating a new MC worksheet.");
    createNewMC(workbook, config_sheet, mc_template_sheet, total_db, individual_db, dates_list, desired_mc_worksheet_name);
  }
  // DEBUG
  // console.log("total_db: ", total_db);
  // console.log("individual_db: ", individual_db);

}

function updateMcWorksheet(worksheet: ExcelScript.Worksheet, total_database: { [key: string]: { [key: string]: number } }, individual_databases: { [key: string]: { [key: string]: object } }, dates_list: string[]){
  console.log("placeholder");
  // Adding/Updating new quantities
  // addNewTotalQuantities(worksheet, total_database);
  // Replace placeholders with dates
  updateDatesStrings(worksheet, dates_list);
  // Adding new total quantities
  // addNewTotalQuantities(worksheet, total_database);
  // Creating list of PO Quantities ranges
  let po_ranges_list = getPoQuantityRanges(worksheet);
  // Adding quantities per periods
  addQuantitiesPerPeriods(worksheet, po_ranges_list, dates_list, individual_databases);
}

function createNewMC(workbook: ExcelScript.Workbook, config_sheet: ExcelScript.Worksheet, mc_template_sheet: ExcelScript.Worksheet, total_database: { [key: string]: { [key: string]: number } }, individual_databases: { [key: string]: { [key: string]: object } }, dates_list: string[], new_mc_tab_name: unknown) {

  // Extracting the remaining quantities from the master sheet tab
  let remaining_quantities_sheet = workbook.getWorksheet("master_sheet");
  let remaining_quantities_values = remaining_quantities_sheet.getRange("E3:E206").getValues();
  // Duplicating template tab
  let new_mc_sheet = mc_template_sheet.copy(ExcelScript.WorksheetPositionType.after, config_sheet);
  new_mc_sheet.setName(new_mc_tab_name.toString());
  // Adding remaining quantities to new mc sheet
  let remaining_quantities_range = new_mc_sheet.getRange("G3:G206");
  remaining_quantities_range.setValues(remaining_quantities_values);
  // Adding new total quantities
  // addNewTotalQuantities(new_mc_sheet, total_database);
  // Replace placeholders with dates
  addDatesStrings(new_mc_sheet, dates_list);
  // Creating list of PO Quantities ranges
  let po_ranges_list = getPoQuantityRanges(new_mc_sheet);
  // Adding quantities per periods
  addQuantitiesPerPeriods(new_mc_sheet, po_ranges_list, dates_list, individual_databases);
}

function getDesiredMcWorksheetName(worksheet: ExcelScript.Worksheet, start_date: unknown){
  // Getting the desired new MC tab name using an XLOOKLUP function implementation
  let po_start_date_cell_address = findCellAddress(worksheet, "PO Start Date");
  let onglet_mc_cell_address = findCellAddress(worksheet, "Onglet MC");
  let lookup_range = findLookupOrReturnRange(worksheet, po_start_date_cell_address);
  let return_range = findLookupOrReturnRange(worksheet, onglet_mc_cell_address);
  let new_mc_tab_name = xLookUp(worksheet, lookup_range, return_range, start_date);
  return new_mc_tab_name;
}

function createDatabase(workbook: ExcelScript.Workbook, bl_tabs_list: string[]): { totalDatabases: { [key: string]: { [key: string]: number } }, individualDatabases: { [key: string]: { [key: string]: object } } }{
  let workpackages_database = {}; // Object used to store the items and total quantities of all BLs
  let individual_databases = {};  // Object used to store the items and total quantities of each BL separately.
  // Iterating through the values of the bl_tabs_list
  for (const current_bl_tab_name of bl_tabs_list) {
    // Initializing the individual tabs database for the current tab
    individual_databases[current_bl_tab_name] = {};
    // Getting the worksheet of the current BL
    let current_worksheet = workbook.getWorksheet(current_bl_tab_name);
    // Getting the addresses of the "item number WP", "Complexity" and "Quantity Ordered Quantité commandée"
    let item_number_wp_address = findCellAddress(current_worksheet, "Item number WP");
    let complexity_address = findCellAddress(current_worksheet, "Complexity");
    let quantities_address = findCellAddress(current_worksheet, "Quantity Ordered Quantité commandée");
    // Getting the values of the "item number WP", "Complexity" and "Quantity Ordered Quantité commandée"
    let item_number_wp_values = findLookupOrReturnRange(current_worksheet, item_number_wp_address).getValues();
    let complexity_values = findLookupOrReturnRange(current_worksheet, complexity_address).getValues();
    let quantities_values = findLookupOrReturnRange(current_worksheet, quantities_address).getValues();
    // Iterating through the values in the "item number WP" column
    for (let i = 0; i < item_number_wp_values.length; i++) {
      if (item_number_wp_values[i][0] !== "") {  // If the cell value is not empty
        if (workpackages_database.hasOwnProperty(item_number_wp_values[i][0])) {  // If the item already exists
          if (workpackages_database[item_number_wp_values[i][0]].hasOwnProperty(complexity_values[i][0])) {  // If complexity already exists
            workpackages_database[item_number_wp_values[i][0]][complexity_values[i][0]] += quantities_values[i][0]; // Adding quantities to existing item and complexity
          } else { // If complexity doesn't exist
            workpackages_database[item_number_wp_values[i][0]][complexity_values[i][0]] = quantities_values[i][0];  // Adding quantities for new complexity
          }
        } else {  // If item doesn't exist
          workpackages_database[item_number_wp_values[i][0]] = {};  // Add new item and create sub-obj for item
          workpackages_database[item_number_wp_values[i][0]][complexity_values[i][0]] = quantities_values[i][0];  // Add new complexity and quantities to item
        }
      }
      // Adding data to individual databases object
      if (item_number_wp_values[i][0] !== "") {  // If the cell value is not empty
        if (individual_databases[current_bl_tab_name].hasOwnProperty(item_number_wp_values[i][0])) {  // If the item already exists
          if (individual_databases[current_bl_tab_name][item_number_wp_values[i][0]].hasOwnProperty(complexity_values[i][0])) {  // If complexity already exists
            individual_databases[current_bl_tab_name][item_number_wp_values[i][0]][complexity_values[i][0]] += quantities_values[i][0]; // Adding quantities to existing item and complexity
          } else { // If complexity doesn't exist
            individual_databases[current_bl_tab_name][item_number_wp_values[i][0]][complexity_values[i][0]] = quantities_values[i][0];  // Adding quantities for new complexity
          }
        } else {  // If item doesn't exist
          individual_databases[current_bl_tab_name][item_number_wp_values[i][0]] = {};  // Add new item and create sub-obj for item
          individual_databases[current_bl_tab_name][item_number_wp_values[i][0]][complexity_values[i][0]] = quantities_values[i][0];  // Add new complexity and quantities to item
        }
      }
    }
  }
  return {totalDatabases: workpackages_database, individualDatabases: individual_databases};
}

function getListByName(config_sheet: ExcelScript.Worksheet, relative_start_and_end_row_indices: {start: number, end: number}, name: string): string[] {
  // Iterating through onglet bl
  let relevant_bl_tabs_list: string[] = [];
  let onglet_bl_cell_address = findCellAddress(config_sheet, name);
  let onglet_bl_values = findLookupOrReturnRange(config_sheet, onglet_bl_cell_address).getValues();
  for (let i = relative_start_and_end_row_indices.start; i <= relative_start_and_end_row_indices.end; i++){
    relevant_bl_tabs_list.push(String(onglet_bl_values[i][0]));
    }
  return relevant_bl_tabs_list
  }

  function getRelevantRelativeStartandEnd(config_sheet: ExcelScript.Worksheet, start_date: unknown, end_date: unknown): { start: number, end: number } {
  // Finding relevant addresses
  let po_start_date_cell_address = findCellAddress(config_sheet, "PO Start Date");
  let po_end_date_cell_address = findCellAddress(config_sheet, "PO End Date");
  let onglet_bl_cell_address = findCellAddress(config_sheet, "Onglet BL");
  // Finding the ranges of the relevant addresses
  let po_start_date_values = findLookupOrReturnRange(config_sheet, po_start_date_cell_address).getValues();
  let po_end_date_values = findLookupOrReturnRange(config_sheet, po_end_date_cell_address).getValues();
  let onglet_bl_values = findLookupOrReturnRange(config_sheet, onglet_bl_cell_address).getValues();
  // Finding the row index of the given start_date and end_date
  let relative_start_row_index = po_start_date_values.findIndex(row => row[0] === start_date);
  let relative_end_row_index = po_end_date_values.findIndex(row => row[0] === end_date);
  return {start: relative_start_row_index, end: relative_end_row_index };
}

function addQuantitiesPerPeriods(worksheet: ExcelScript.Worksheet, po_ranges_list: ExcelScript.Range[], dates_list: string[], individual_databases: { [key: string]: { [key: string]: object } }){
  // Getting the workpackages range values
  let wp_values = worksheet.getRange("B3:B206").getValues();
  // Iterating through each sub-database
  let iter = 0;
  for (const bl_tab_name in individual_databases){
    let current_database = individual_databases[bl_tab_name];
    let current_po_range = po_ranges_list[iter++];
    // Iterate though the items (workpackage names) in the database and find them in the newly created worksheet
    for (const item in current_database) {
      let item_found = false;  // boolean used to verify that each item in database was found in the worksheet
      let condition_met = false;  // boolean used to verify that data was added to worksheet for corresponding item
      for (let i = 0; i < wp_values.length; i++) {
        if (item === wp_values[i][0]) {
          item_found = true;

          const item_data = current_database[item];
          if (item_data.hasOwnProperty("No complexity")) {
            current_po_range.getCell(i, 0).setValue(item_data["No complexity"]);
            condition_met = true;
          }
          if (item_data.hasOwnProperty("High")) {
            current_po_range.getCell(i, 0).setValue(item_data["High"]);
            condition_met = true;
          }
          if (item_data.hasOwnProperty("Medium")) {
            current_po_range.getCell(i + 1, 0).setValue(item_data["Medium"]);
            condition_met = true;
          }
          if (item_data.hasOwnProperty("Low")) {
            current_po_range.getCell(i + 2, 0).setValue(item_data["Low"]);
            condition_met = true;
          }
          if (!condition_met) {  // Verifying if data was added to newly created worksheet
            // If the flag is still false, none of the conditions were true
            throw new Error(`No complexity, High, Medium, or Low properties found for item: ${item}`);
          }
          break;
        }
      }
      if (!item_found) {  // Verifying if current item in database was found in newly created worksheet
        // If the flag is still false, none of the conditions were true
        throw new Error(`Item not found in column B: ${item}`);
      }
    }
  }
}

// Function that returns the 5 PO quantities ranges for a given worksheet
function getPoQuantityRanges(worksheet: ExcelScript.Worksheet): ExcelScript.Range[]{
  let po1_range = worksheet.getRange("L3:L206");
  let po2_range = worksheet.getRange("P3:P206");
  let po3_range = worksheet.getRange("T3:T206");
  let po4_range = worksheet.getRange("X3:X206");
  let po5_range = worksheet.getRange("AB3:AB206");
  return [po1_range, po2_range, po3_range, po4_range, po5_range];
}

function updateDatesStrings(worksheet: ExcelScript.Worksheet, date_list: string[]) {
  // Getting the placeholders range values
  let placeholder_range = worksheet.getRange("L1:AE1");
  let placeholder_range_values = placeholder_range.getValues()[0];  // the [0] is necessary to since we would get an array of arrays otherwise
  // Perform first iteration to see if any new dates were updates (mainly, to find the iterator index)
  let start_index = 0;
  for (let i = 0; i < date_list.length; i++) {
    for (let j = 0; j < placeholder_range_values.length; j++) {
      if (placeholder_range_values[j] === date_list[i]){ // If date_list[i] exists in the placeholder range
        start_index++; // Incrementing starting index
        break;
      }
    }
  }
  // Replacing the placeholders with the dates from date_list (Starting from start_index)
  for (let i = start_index; i < date_list.length; i++) {
    for (let j = 0; j < placeholder_range_values.length; j++) {
      if (placeholder_range_values[j] === "PO" + (i + 1)) {
        placeholder_range.getCell(0, j).setValue(date_list[i]);
        break;
      }
    }
  }
}

function addDatesStrings(worksheet: ExcelScript.Worksheet, date_list: string[]){
  // Getting the placeholders range values
  let placeholder_range = worksheet.getRange("L1:AE1");
  let placeholder_range_values = placeholder_range.getValues()[0];  // the [0] is necessary to since we would get an array of arrays otherwise
  // Replacing the placeholders with the dates from date_list
  for (let i=0; i<date_list.length; i++){
    for (let j = 0; j < placeholder_range_values.length; j++){
      if (placeholder_range_values[j] === "PO"+(i+1)){
        placeholder_range.getCell(0, j).setValue(date_list[i]);
        break;
      }
    }
  }
}

function addNewTotalQuantities(worksheet: ExcelScript.Worksheet, database: { [key: string]: { [key: string]: number } }){
  // Get relevant ranges
  let wp_values = worksheet.getRange("B3:B206").getValues();
  let quantities_ranges = worksheet.getRange("F3:F206");
  // Iterate though the items (workpackage names) in the database and find them in the newly created worksheet
  for (const item in database){
    let item_found = false;  // boolean used to verify that each item in database was found in the worksheet
    let condition_met = false;  // boolean used to verify that data was added to worksheet for corresponding item
    for (let i = 0; i < wp_values.length; i++){
      if (item === wp_values[i][0]){
        item_found = true;
        const item_data = database[item];

        if (item_data.hasOwnProperty("No complexity")) {
          quantities_ranges.getCell(i, 0).setValue(item_data["No complexity"]);
          condition_met = true;
        }
        if (item_data.hasOwnProperty("High")){
          quantities_ranges.getCell(i, 0).setValue(item_data["High"]);
          condition_met = true;
        }
        if (item_data.hasOwnProperty("Medium")) {
          quantities_ranges.getCell(i + 1, 0).setValue(item_data["Medium"]);
          condition_met = true;
        }
        if (item_data.hasOwnProperty("Low")) {
          quantities_ranges.getCell(i + 2, 0).setValue(item_data["Low"]);
          condition_met = true;
        }
        if (!condition_met) {  // Verifying if data was added to newly created worksheet
          // If the flag is still false, none of the conditions were true
          throw new Error(`No complexity, High, Medium, or Low properties found for item: ${item}`);
        }
        break; // skipping current item if found
      }
    }
    if (!item_found){  // Verifying if current item in database was found in newly created worksheet
      // If the flag is still false, none of the conditions were true
      throw new Error(`Item not found in column B: ${item}`);
    }
  }
}

function getTotalRowsCount(worksheet: ExcelScript.Worksheet): number {
  let worksheet_used_range = worksheet.getUsedRange();
  return worksheet_used_range.getRowCount();
}

  function findCellIndices(worksheet: ExcelScript.Worksheet, searchText: string): {rowIndex: number, colIndex: number} {
  let range = worksheet.getUsedRange();
  let found_range = range.find(searchText, { completeMatch: true });
  return { rowIndex: found_range.getRowIndex(), colIndex: found_range.getColumnIndex() };
}

function findCellAddress(worksheet: ExcelScript.Worksheet, searchText: string): string {
  let range = worksheet.getUsedRange();
  let found_cell = range.find(searchText, { completeMatch: true });
  return found_cell.getAddress();
}

// Function used to get the range from a found cell address to the final used row in the worksheet
function findLookupOrReturnRange(worksheet: ExcelScript.Worksheet, found_title_cell_address: string):ExcelScript.Range {
  let found_cell_range = worksheet.getRange(found_title_cell_address);
  // Getting the total number of rows used in the worksheet
  let total_used_rows = worksheet.getUsedRange().getRowCount();
  // Calculating the row index for the cell just below the found_title_cell
  let start_row_index = found_cell_range.getRowIndex() + 1;
  // Calculating the number of rows to include in the range from the cell below to the end
  let row_count = total_used_rows - start_row_index;
  // Getting the range from the cell just below the found title cell to the end of the used range
  let desired_range = worksheet.getRangeByIndexes(start_row_index, found_cell_range.getColumnIndex(), row_count, 1);
  return desired_range;
}

// Function that implements the XLOOKUP Excel function
function xLookUp(worksheet: ExcelScript.Worksheet, lookup_range: ExcelScript.Range, return_range: ExcelScript.Range, lookup_value: unknown):unknown {
  // Getting the values from the lookup range
  let lookup_values = lookup_range.getValues();
  // Getting the index of the lookup value
  let row_index = lookup_values.findIndex(row => row[0] === lookup_value);
  // Checking if the desired value was found
  // If the value is found, get the corresponding return value
  let return_value = return_range.getCell(row_index, 0).getValue();
  return return_value;
}
