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
  
  console.log(total_db);
  // Creating New MC using the MC tab name at the desired template
  createNewMC(workbook, config_sheet, mc_template_sheet, start_date, total_db, individual_db);
}

function createDatabase(workbook: ExcelScript.Workbook, bl_tabs_list: string[]): { totalDatabases: object, individualDatabases: object }{
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

function getListByName(config_sheet: ExcelScript.Worksheet, relative_start_and_end_row_indices: object, name: string): string[] {
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

function createNewMC(workbook: ExcelScript.Workbook, config_sheet: ExcelScript.Worksheet, mc_template_sheet: ExcelScript.Worksheet, start_date: unknown, total_database: object, individual_databases: object) {
  // Getting the desired new MC tab name using an XLOOKLUP function implementation
  let po_start_date_cell_address = findCellAddress(config_sheet, "PO Start Date");
  let onglet_mc_cell_address = findCellAddress(config_sheet, "Onglet MC");
  let lookup_range = findLookupOrReturnRange(config_sheet, po_start_date_cell_address);
  let return_range = findLookupOrReturnRange(config_sheet, onglet_mc_cell_address);
  let new_mc_tab_name = xLookUp(config_sheet, lookup_range, return_range, start_date);
  // Extracting the remaining quantities from the master sheet tab
  let remaining_quantities_sheet = workbook.getWorksheet("master_sheet");
  let remaining_quantities_values = remaining_quantities_sheet.getRange("E3:E206").getValues();
  // Duplicating template tab
  let new_mc_sheet = mc_template_sheet.copy(ExcelScript.WorksheetPositionType.after, config_sheet);
  new_mc_sheet.setName(new_mc_tab_name.toString());
  // Adding remaining quantities to new mc sheet
  let remaining_quantities_range = new_mc_sheet.getRange("G3:G206");
  remaining_quantities_range.setValues(remaining_quantities_values);
  // Adding new quantities
  addNewTotalQuantities(new_mc_sheet, total_database);
}

function addNewTotalQuantities(worksheet: ExcelScript.Worksheet, database: object){
  // Get relevant ranges
  let wp_values = worksheet.getRange("B3:B206").getValues();
  let quantities_ranges = worksheet.getRange("E3:E206");
  console.log("database= ", database);
  // Iterate though the ranges of the item with the object and check if item exist and correspond to current item row
  for (let i = 0; i < wp_values.length; i++){
    let current_name = wp_values[i][0];
    if (database.hasOwnProperty(current_name)){
      console.log(current_name);
    }
  }
  // If item exists, and if complexity is not "No complexity", then iterate thorugh the complexities High, Medium, Low (or nest if and add value relative to current row (+0 ,+1, or +2))
  // If item exists, but complexity is No complexity, add directly to row + 0
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
