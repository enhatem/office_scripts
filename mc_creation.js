
function main(workbook: ExcelScript.Workbook) {

  // Getting the worksheets that will be used to duplicate the script 
  let sheet = workbook.getActiveWorksheet();
  let config_sheet = workbook.getWorksheet("config");
  let mc_template_sheet = workbook.getWorksheet("MC_Template")

  // Getting the Start and End Date of the current period
  let start_date = config_sheet.getRange("C6").getValue();
  let end_date = config_sheet.getRange("C7").getValue();
  
  // Creating New MC using the MC tab name at the desired template
  createNewMC(workbook, config_sheet, mc_template_sheet, start_date);

}

function createNewMC(workbook: ExcelScript.Workbook, config_sheet: ExcelScript.Worksheet, mc_template_sheet: ExcelScript.Worksheet, start_date: unknown){

  // Getting the desired new MC tab name using an XLOOKLUP function implementation
  let po_start_date_cell_address = findCellAddress(config_sheet, "PO Start Date");
  let onglet_mc_cell_address = findCellAddress(config_sheet, "Onglet MC");
  let lookup_range = findLookupOrReturnRange(config_sheet, po_start_date_cell_address);
  let return_range = findLookupOrReturnRange(config_sheet, onglet_mc_cell_address);
  let new_mc_tab_name = xLookUp(config_sheet, lookup_range, return_range, start_date);
  console.log(new_mc_tab_name);
  console.log(typeof new_mc_tab_name);

  // Extracting the remaining quantities from the master sheet tab
  let remaining_quantities_sheet = workbook.getWorksheet("master_sheet");
  let remaining_quantities_values = remaining_quantities_sheet.getRange("E3:E206").getValues();

  // Duplicating template tab
  let new_mc_sheet = mc_template_sheet.copy(ExcelScript.WorksheetPositionType.after, config_sheet);
  new_mc_sheet.setName(new_mc_tab_name.toString());

  // Adding remaining quantities to new mc sheet
  let remaining_quantities_range = new_mc_sheet.getRange("G3:G206");
  remaining_quantities_range.setValues(remaining_quantities_values);
}

function getTotalRowsCount(worksheet: ExcelScript.Worksheet) {
  let worksheet_used_range = worksheet.getUsedRange();
  return worksheet_used_range.getRowCount(); 
}

function findCellIndices(worksheet: ExcelScript.Worksheet, searchText: string){
  let range = worksheet.getUsedRange();
  let found_range = range.find(searchText, { completeMatch: true });
  return {rowIndex: found_range.getRowIndex(), colIndex: found_range.getColumnIndex()};
}

function findCellAddress(worksheet: ExcelScript.Worksheet, searchText: string) {
  let range = worksheet.getUsedRange();
  let found_cell = range.find(searchText, { completeMatch: true });
  return found_cell.getAddress();
}

// Function used to get the range from a found cell address to the final used row in the worksheet
function findLookupOrReturnRange(worksheet: ExcelScript.Worksheet, found_title_cell_address: string){
  let found_cell_range = worksheet.getRange(found_title_cell_address);
  // Getting the total number of rows used in the worksheet
  let total_used_rows = worksheet.getUsedRange().getRowCount();
  // Calculating the row index for the cell just below the found_title_cell
  let start_row_index = found_cell_range.getRowIndex() + 1;
  // Calculating the number of rows to include in the range from the cell below to the end
  let row_count = total_used_rows - start_row_index;
  // Getting the range from the cell just below the found title cell to the end of the used range
  let desired_range = worksheet.getRangeByIndexes(start_row_index,found_cell_range.getColumnIndex(), row_count, 1);
  return desired_range;
}

// Function that implements the XLOOKUP Excel function
function xLookUp(worksheet: ExcelScript.Worksheet, lookup_range: ExcelScript.Range, return_range: ExcelScript.Range, lookup_value: unknown)
{
  // Getting the values from the lookup range
  let lookup_values = lookup_range.getValues();
  // Getting the index of the lookup value
  let row_index = lookup_values.findIndex(row => row[0] === lookup_value);
  // Checking if the desired value was found
  // If the value is found, get the corresponding return value
  let return_value = return_range.getCell(row_index,0).getValue();
  return return_value;
}
