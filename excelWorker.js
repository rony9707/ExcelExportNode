const { parentPort } = require('worker_threads');
const ExcelJS = require('exceljs');

/**
 * Function to generate an Excel file buffer.
 * @param {Array} data - Array of objects representing the data rows.
 * @param {Array} config - Array of objects representing the column config.
 * @returns {Promise<Buffer>}
 */
async function generateExcelBuffer(data, config) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Employees');

    // Set headers
    worksheet.columns = config.map(col => ({
        header: col.label,
        key: col.key,
        width: Math.max(col.label.length + 2, 10)
    }));

    // Style headers
    const headerRow = worksheet.getRow(1);
    headerRow.eachCell(cell => {
        cell.font = { bold: true };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCFFCC' } };
    });

    // Add data rows
    data.forEach(row => worksheet.addRow(row));

    // Adjust column widths
    worksheet.columns.forEach(column => {
        let maxLength = column.width || 10;
        data.forEach(row => {
            const value = row[column.key];
            if (value) {
                maxLength = Math.max(maxLength, value.toString().length);
            }
        });
        column.width = Math.min(maxLength + 2, 100);
    });

    // Generate buffer and send it back
    const buffer = await workbook.xlsx.writeBuffer();
    parentPort.postMessage(buffer);
}

// Listen for messages from the main thread
parentPort.on('message', async ({ data, config }) => {
    try {
        const buffer = await generateExcelBuffer(data, config);
        parentPort.postMessage(buffer);
    } catch (error) {
        parentPort.postMessage({ error: error.message });
    }
});
