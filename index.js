const ExcelJS = require('exceljs');

// Load the existing workbook or create a new one
const workbook = new ExcelJS.Workbook();

// Load the existing worksheet or create a new one
const sheet = workbook.addWorksheet('Sheet1');

// Function to save the workbook to a file
const saveWorkbook = async () => {
  try {
    await workbook.xlsx.writeFile('data.xlsx');
    console.log('Workbook saved successfully.');
  } catch (error) {
    console.error('Failed to save the workbook:', error);
  }
};

// Function to create a new entry in the Excel sheet
const createEntry = async (data) => {
  try {
    // Here, you need to define the logic to add the data to the Excel sheet.
    // For example, you can add the data to the next available row in the worksheet.

    // Assuming data is an object with fields that you want to save.
    const newRow = sheet.addRow(data);
    await saveWorkbook();
    console.log('Entry created successfully.');
  } catch (error) {
    console.error('Failed to create entry:', error);
  }
};

// Function to read data from the Excel sheet
const readData = async () => {
  try {
    // Here, you need to define the logic to read data from the Excel sheet.
    // For example, you can read all rows from the worksheet and return them as an array of objects.

    const rows = sheet.getRows();
    const data = rows.map((row) => row.values);
    console.log('Data read successfully:', data);
    return data;
  } catch (error) {
    console.error('Failed to read data:', error);
    return [];
  }
};

// Function to update an entry in the Excel sheet
const updateEntry = async (rowIndex, newData) => {
  try {
    // Here, you need to define the logic to update the data in the Excel sheet based on the rowIndex.
    // For example, you can find the row with the given index and update its values with the newData.

    const row = sheet.getRow(rowIndex);
    Object.keys(newData).forEach((key) => {
      row.getCell(key).value = newData[key];
    });
    await saveWorkbook();
    console.log('Entry updated successfully.');
  } catch (error) {
    console.error('Failed to update entry:', error);
  }
};

// Function to delete an entry from the Excel sheet
const deleteEntry = async (rowIndex) => {
  try {
    // Here, you need to define the logic to delete the data from the Excel sheet based on the rowIndex.
    // For example, you can find the row with the given index and remove it from the worksheet.

    sheet.spliceRows(rowIndex, 1);
    await saveWorkbook();
    console.log('Entry deleted successfully.');
  } catch (error) {
    console.error('Failed to delete entry:', error);
  }
};

// Example usage:

// To create a new entry:
createEntry({ field1: 'Value 1', field2: 'Value 2' });

// To read data from the Excel sheet:
readData();

// To update an entry (e.g., row 2):
updateEntry(2, { field1: 'Updated Value 1' });

// To delete an entry (e.g., row 3):
deleteEntry(3);
