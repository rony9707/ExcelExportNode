const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const { Worker } = require('worker_threads');
const app = express();
const port = 3000;


app.use(bodyParser.json({ limit: '10mb' }));

app.post('/download-excel', async (req, res) => {
    const { data, config } = req.body;

    if (!Array.isArray(data) || data.length === 0) {
        return res.status(400).send('Invalid data');
    }

    // Use default config if none is provided
    const finalConfig = config && Array.isArray(config) && config.length > 0
        ? config
        : generateConfigFromData(data);

    // Create a new worker thread
    const worker = new Worker('./excelWorker.js');

    // Send data to worker thread
    worker.postMessage({ data, config: finalConfig });

    // Listen for messages from worker thread
    worker.on('message', (buffer) => {
        res.setHeader('Content-Disposition', 'attachment; filename="data.xlsx"');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
    });

    worker.on('error', (error) => {
        console.error('Worker error:', error);
        res.status(500).send('Error generating Excel file');
    });

    worker.on('exit', (code) => {
        if (code !== 0) {
            console.error(`Worker stopped with exit code ${code}`);
        }
    });
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
        type: 'text' // Default type
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
