const XLSX = require('xlsx');
const fs = require('fs');
const https = require('https');

class ExcelReader {
    constructor(filePath) {
        this.filePath = filePath;
        this.workbook = null;
        this.sheetNames = [];
        this.isGoogleSheets = this.filePath.includes('docs.google.com/spreadsheets');
        this.cachedData = null;
        this.cacheTime = null;
        this.cacheTimeout = 5 * 60 * 1000; // 5 minutes cache
    }

    /**
     * Convert Google Sheets URL to CSV export URL
     * @param {string} url - Google Sheets URL
     * @returns {string} CSV export URL
     */
    convertToCSVUrl(url) {
        const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
        if (!match) {
            throw new Error('Invalid Google Sheets URL');
        }
        const spreadsheetId = match[1];

        // Extract gid if present
        const gidMatch = url.match(/gid=([0-9]+)/);
        const gid = gidMatch ? gidMatch[1] : '0';

        return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&gid=${gid}`;
    }

    /**
     * Fetch CSV data from URL
     * @param {string} url - CSV URL
     * @returns {Promise<string>} CSV data
     */
    fetchCSV(url) {
        return new Promise((resolve, reject) => {
            https.get(url, (response) => {
                // Handle redirects
                if (response.statusCode === 301 || response.statusCode === 302 || response.statusCode === 307) {
                    return this.fetchCSV(response.headers.location).then(resolve).catch(reject);
                }

                if (response.statusCode !== 200) {
                    reject(new Error(`Failed to fetch CSV: ${response.statusCode}`));
                    return;
                }

                let data = '';
                response.on('data', chunk => data += chunk);
                response.on('end', () => resolve(data));
            }).on('error', reject);
        });
    }

    /**
     * Parse CSV to array of objects
     * @param {string} csv - CSV string
     * @returns {Array} Parsed data
     */
    parseCSV(csv) {
        const lines = csv.trim().split('\n');
        if (lines.length === 0) return [];

        const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const line = lines[i];
            if (!line.trim()) continue;

            const values = [];
            let current = '';
            let inQuotes = false;

            // Handle CSV parsing with quoted values
            for (let j = 0; j < line.length; j++) {
                const char = line[j];
                if (char === '"') {
                    inQuotes = !inQuotes;
                } else if (char === ',' && !inQuotes) {
                    values.push(current.trim().replace(/^"|"$/g, ''));
                    current = '';
                } else {
                    current += char;
                }
            }
            values.push(current.trim().replace(/^"|"$/g, ''));

            // Create object from headers and values
            const row = {};
            let hasData = false;
            headers.forEach((header, index) => {
                const value = values[index] || '';
                row[header] = value;
                // Check if row has any non-zero, non-empty data
                if (value && value !== '0' && value !== '0.0') {
                    hasData = true;
                }
            });

            // Only add rows with actual data
            if (hasData) {
                data.push(row);
            }
        }

        return data;
    }

    /**
     * Load data from Google Sheets
     * @returns {Promise<boolean>} Success status
     */
    async loadGoogleSheets() {
        try {
            // Check cache
            if (this.cachedData && this.cacheTime && (Date.now() - this.cacheTime) < this.cacheTimeout) {
                return true;
            }

            const csvUrl = this.convertToCSVUrl(this.filePath);
            const csvData = await this.fetchCSV(csvUrl);

            this.cachedData = this.parseCSV(csvData);
            this.cacheTime = Date.now();
            this.sheetNames = ['Sheet1']; // Google Sheets CSV export is single sheet

            return true;
        } catch (error) {
            console.error('[ERROR] Failed to load Google Sheets:', error.message);
            return false;
        }
    }

    /**
     * Load the Excel file or Google Sheets
     * @returns {Promise<boolean>|boolean} Success status
     */
    loadFile() {
        // If it's a Google Sheets URL, use async loading
        if (this.isGoogleSheets) {
            return this.loadGoogleSheets();
        }

        // Otherwise load local Excel file
        try {
            if (!fs.existsSync(this.filePath)) {
                throw new Error(`Excel file not found at: ${this.filePath}`);
            }

            this.workbook = XLSX.readFile(this.filePath);
            this.sheetNames = this.workbook.SheetNames;
            return true;
        } catch (error) {
            console.error('[ERROR] Failed to load Excel file:', error.message);
            return false;
        }
    }

    /**
     * Get all sheet names in the workbook
     * @returns {Promise<string[]>|string[]} Array of sheet names
     */
    async getSheetNames() {
        if (this.isGoogleSheets) {
            await this.loadFile();
            return this.sheetNames;
        }

        if (!this.workbook) {
            this.loadFile();
        }
        return this.sheetNames;
    }

    /**
     * Get data from a specific sheet
     * @param {string} sheetName - Name of the sheet (optional, defaults to first sheet)
     * @returns {Promise<Array>|Array} Sheet data as array of objects
     */
    async getSheetData(sheetName = null) {
        if (this.isGoogleSheets) {
            await this.loadFile();
            return this.cachedData || [];
        }

        if (!this.workbook) {
            this.loadFile();
        }

        const targetSheet = sheetName || this.sheetNames[0];

        if (!this.workbook.Sheets[targetSheet]) {
            throw new Error(`Sheet "${targetSheet}" not found`);
        }

        const sheet = this.workbook.Sheets[targetSheet];
        return XLSX.utils.sheet_to_json(sheet);
    }

    /**
     * Get all data from all sheets
     * @returns {Promise<Object>|Object} Object with sheet names as keys and data arrays as values
     */
    async getAllData() {
        if (this.isGoogleSheets) {
            await this.loadFile();
            return { 'Sheet1': this.cachedData || [] };
        }

        if (!this.workbook) {
            this.loadFile();
        }

        const allData = {};
        this.sheetNames.forEach(sheetName => {
            allData[sheetName] = this.getSheetData(sheetName);
        });

        return allData;
    }

    /**
     * Search for data containing a specific value
     * @param {string} searchTerm - Term to search for
     * @param {string} sheetName - Sheet to search in (optional)
     * @returns {Promise<Array>|Array} Array of matching rows
     */
    async search(searchTerm, sheetName = null) {
        const data = await this.getSheetData(sheetName);
        const lowerSearchTerm = searchTerm.toLowerCase();

        return data.filter(row => {
            return Object.values(row).some(value => {
                return String(value).toLowerCase().includes(lowerSearchTerm);
            });
        });
    }

    /**
     * Get statistics for a numeric column
     * @param {string} columnName - Name of the column
     * @param {string} sheetName - Sheet name (optional)
     * @returns {Promise<Object>|Object} Statistics object
     */
    async getColumnStats(columnName, sheetName = null) {
        const data = await this.getSheetData(sheetName);

        const values = data
            .map(row => row[columnName])
            .filter(val => val !== undefined && val !== null && !isNaN(val))
            .map(val => Number(val));

        if (values.length === 0) {
            return { error: `No numeric data found for column "${columnName}"` };
        }

        const sum = values.reduce((acc, val) => acc + val, 0);
        const avg = sum / values.length;
        const min = Math.min(...values);
        const max = Math.max(...values);

        // Calculate median
        const sorted = [...values].sort((a, b) => a - b);
        const mid = Math.floor(sorted.length / 2);
        const median = sorted.length % 2 === 0
            ? (sorted[mid - 1] + sorted[mid]) / 2
            : sorted[mid];

        return {
            count: values.length,
            sum: sum,
            average: avg,
            min: min,
            max: max,
            median: median,
        };
    }

    /**
     * Get row count
     * @param {string} sheetName - Sheet name (optional)
     * @returns {Promise<number>|number} Number of rows
     */
    async getRowCount(sheetName = null) {
        const data = await this.getSheetData(sheetName);
        return data.length;
    }

    /**
     * Get column names (headers)
     * @param {string} sheetName - Sheet name (optional)
     * @returns {Promise<string[]>|string[]} Array of column names
     */
    async getColumns(sheetName = null) {
        const data = await this.getSheetData(sheetName);
        if (data.length === 0) return [];

        return Object.keys(data[0]);
    }
}

module.exports = ExcelReader;
