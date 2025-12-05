require('dotenv').config();

module.exports = {
    // Discord bot token from .env file
    token: process.env.DISCORD_TOKEN,

    // Client ID for slash commands
    clientId: process.env.CLIENT_ID,

    // Excel file path or Google Sheets URL
    excelFilePath: process.env.EXCEL_FILE_PATH || process.env.GOOGLE_SHEETS_URL || './data/sample.xlsx',

    // Bot prefix for text commands (optional)
    prefix: process.env.PREFIX || '!',
};
