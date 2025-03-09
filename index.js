const express = require('express');
const ExcelJS = require('exceljs');
const cors = require('cors');
const bodyParser = require('body-parser');

const app = express();
const port = 3000;

app.use(cors());

//Limit 10MB of data
app.use(bodyParser.json({ limit: '10mb' }));

app.post('/download-excel', async (req, res) => {
  const { data, config } = req.body;

  if (!Array.isArray(data) || data.length === 0) {
      return res.status(400).send('Invalid data');
  }

  // Check if config is provided, if not, generate it from the data keys
  const finalConfig = config && Array.isArray(config) && config.length > 0
      ? config
      : generateConfigFromData(data);

  try {
      const buffer = await generateExcelBuffer(data, finalConfig);

      res.setHeader('Content-Disposition', 'attachment; filename="data.xlsx"');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(buffer);
  } catch (error) {
      console.error('Error generating Excel:', error);
      res.status(500).send('Failed to generate Excel file');
  }
});



/**
* Generates column configuration from the first object in data.
* @param {Array} data
* @returns {Array}
*/
function generateConfigFromData(data) {
  const firstRow = data[0];
  return Object.keys(firstRow).map(key => ({
      key: key,
      label: capitalizeFirstLetter(key),
      type: 'text'  // Default type for simplicity, but you can extend this logic
  }));
}

/**
* Capitalizes the first letter of a string.
* @param {string} str
* @returns {string}
*/
function capitalizeFirstLetter(str) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}


app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});

/**
 * Generates an Excel file buffer from data and config.
 * @param {Array} data - Array of objects representing the data rows.
 * @param {Array} config - Array of objects representing the column config (key, label).
 * @returns {Promise<Buffer>}
 */
async function generateExcelBuffer(data, config) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Employees');

    // Set headers and auto filter
    const headers = prepareHeaders(config);
    worksheet.columns = headers;
    setAutoFilter(worksheet, headers.length);

    // Style headers
    styleHeaderRow(worksheet);

    // Add data rows
    addDataRows(worksheet, data, config);

    // Adjust column widths
    adjustColumnWidths(worksheet, data);

    // Generate buffer
    return await workbook.xlsx.writeBuffer();
}


/**
 * Prepares column headers from config.
 * @param {Array} config
 * @returns {Array}
 */
function prepareHeaders(config) {
    return config.map(col => ({
        header: col.label,
        key: col.key,
        width: Math.max(col.label.length + 2, 10)  // Minimum width
    }));
}

/**
 * Sets auto filter on the header row.
 * @param {ExcelJS.Worksheet} worksheet
 * @param {number} columnCount
 */
function setAutoFilter(worksheet, columnCount) {
    const lastColumn = String.fromCharCode(64 + columnCount);  // A, B, C...
    worksheet.autoFilter = `A1:${lastColumn}1`;
}

/**
 * Styles the header row (bold, centered).
 * @param {ExcelJS.Worksheet} worksheet
 */
function styleHeaderRow(worksheet) {
  const headerRow = worksheet.getRow(1);
  headerRow.eachCell(cell => {
      cell.font = { bold: true };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFCCFFCC' } // Light green color
      };
  });
}


/**
 * Adds data rows to the worksheet and appends subtotal for summable columns.
 * @param {ExcelJS.Worksheet} worksheet
 * @param {Array} data
 * @param {Array} config
 */
function addDataRows(worksheet, data, config) {
  let lastRowNumber = 1; // Start from header row index

  data.forEach(row => {
      const newRow = {};
      config.forEach(col => {
          newRow[col.key] = row[col.key];
      });
      worksheet.addRow(newRow);
  });

  lastRowNumber = worksheet.rowCount; // Last data row index

  // Add subtotal row if there are summable columns
  const subtotalRow = {};
  config.forEach(col => {
      if (col.summable) {
          const columnLetter = getColumnLetter(worksheet, col.key);
          subtotalRow[col.key] = { formula: `SUBTOTAL(9, ${columnLetter}2:${columnLetter}${lastRowNumber})` };
      } else {
          subtotalRow[col.key] = '';
      }
  });

  const subtotalExcelRow = worksheet.addRow(subtotalRow);
  styleSubtotalRow(subtotalExcelRow);
}

/**
 * Retrieves the column letter based on the key in the worksheet.
 * @param {ExcelJS.Worksheet} worksheet
 * @param {string} key
 * @returns {string}
 */
function getColumnLetter(worksheet, key) {
  const columnIndex = worksheet.columns.findIndex(col => col.key === key);
  return columnIndex >= 0 ? String.fromCharCode(65 + columnIndex) : '';
}


/**
 * Styles the subtotal row.
 * @param {ExcelJS.Row} row
 */
function styleSubtotalRow(row) {
  row.eachCell(cell => {
      cell.font = { bold: true };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFCC00' } // Light yellow color
      };
  });
}

/**
 * Dynamically adjusts column widths based on data.
 * @param {ExcelJS.Worksheet} worksheet
 * @param {Array} data
 */
function adjustColumnWidths(worksheet, data) {
  worksheet.columns.forEach(column => {
      let maxLength = column.width || 10;

      data.forEach(row => {
          const value = row[column.key];
          if (value) {
              const length = value.toString().length;
              if (length > maxLength) maxLength = length;
          }
      });

      // Cap the max length to 100
      maxLength = Math.min(maxLength, 100);

      column.width = maxLength + 2;  // Padding
  });
}

