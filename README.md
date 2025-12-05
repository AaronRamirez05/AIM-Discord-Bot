# AIM Discord Bot

A Discord bot that reads and reports powerlifting data from Google Sheets (or Excel files). This bot allows you to view, search, generate statistics, and announce rankings directly through Discord commands.

## Features

- Read and display data from Google Sheets (or local Excel files)
- Announce powerlifting rankings with @everyone mention
- Search through lifter data
- Generate statistics and reports
- Support for multiple sheets
- Both slash commands and text commands
- Column-level statistics (sum, average, min, max, median)
- Auto-updating from Google Sheets (5-minute cache)
- Ranked leaderboards with medals for top performers

## Prerequisites

- Node.js (v16.9.0 or higher)
- A Discord Bot Token
- A Google Sheets URL (publicly viewable) OR a local Excel file (.xlsx or .xls)

## Installation

1. Clone this repository:
```bash
git clone https://github.com/AaronRamirez05/AIM-Discord-Bot.git
cd AIM-Discord-Bot
```

2. Install dependencies:
```bash
npm install
```

3. Create a `.env` file by copying `.env.example`:
```bash
cp .env.example .env
```

4. Configure your `.env` file with your Discord bot credentials:
```
DISCORD_TOKEN=your_bot_token_here
CLIENT_ID=your_client_id_here
EXCEL_FILE_PATH=./data/sample.xlsx
PREFIX=!
```

## Setting Up Your Discord Bot

1. Go to the [Discord Developer Portal](https://discord.com/developers/applications)
2. Click "New Application" and give it a name
3. Go to the "Bot" section and click "Add Bot"
4. Copy the bot token and paste it into your `.env` file as `DISCORD_TOKEN`
5. Copy the Application ID from the "General Information" section and paste it as `CLIENT_ID`
6. Under "Bot" section, enable these Privileged Gateway Intents:
   - Message Content Intent
   - Server Members Intent (optional)
7. Go to "OAuth2" > "URL Generator"
8. Select scopes: `bot` and `applications.commands`
9. Select bot permissions:
   - Send Messages
   - Embed Links
   - Read Message History
   - Use Slash Commands
10. Copy the generated URL and use it to invite the bot to your server

## Preparing Your Excel File

1. Place your Excel file in the `data/` directory
2. Ensure your Excel file has headers in the first row
3. Update `EXCEL_FILE_PATH` in `.env` to point to your file

Example Excel structure:
```
| Name  | Age | Score | Department |
|-------|-----|-------|------------|
| John  | 25  | 85    | Sales      |
| Jane  | 30  | 92    | Marketing  |
| Bob   | 28  | 78    | IT         |
```

## Deploying Slash Commands

Before running the bot for the first time, deploy the slash commands:

```bash
node deploy-commands.js
```

This registers all slash commands with Discord. You only need to do this once, or whenever you add/modify commands.

## Running the Bot

Start the bot:
```bash
npm start
```

You should see:
```
[INFO] Loaded command: data
[INFO] Loaded command: search
[INFO] Loaded command: sheets
[INFO] Loaded command: stats
[SUCCESS] Logged in as YourBot#1234
[INFO] Bot is ready and running!
```

## Available Commands

### Slash Commands

1. `/announce-rankings`
   - Post a complete rankings announcement with @everyone mention
   - Shows overview statistics (participants, averages, records)
   - Displays all lifters sorted by Total Dots score
   - Includes medals for top 3 performers

2. `/data [sheet] [limit]`
   - Display data from the Excel sheet
   - `sheet`: (Optional) Sheet name to read from
   - `limit`: (Optional) Number of rows to display (default: 10)

3. `/search <query> [sheet]`
   - Search for data in the Excel sheet
   - `query`: The search term (required)
   - `sheet`: (Optional) Sheet name to search in

4. `/sheets`
   - List all available sheets in the Excel file

5. `/stats overview [sheet]`
   - Get overview statistics of the sheet
   - Shows total rows, columns, column names, and available sheets

6. `/stats column <name> [sheet]`
   - Get statistics for a specific numeric column
   - Shows count, sum, average, min, max, and median

### Text Commands (with prefix)

All slash commands also work as text commands with the prefix `!`:

- `!announce-rankings` - Post rankings announcement
- `!data [limit] [sheet]` - Display data
- `!search <query> [sheet]` - Search data
- `!sheets` - List sheets
- `!stats overview [sheet]` - Overview statistics
- `!stats column <name> [sheet]` - Column statistics

## Examples

### Announce rankings completion
```
/announce-rankings
```
This will:
- Post @everyone mention
- Show overview stats (total participants, averages, records)
- Display all lifters ranked by Total Dots
- Award medals to top 3 performers

### View data from the first sheet
```
/data
```

### View 5 rows from a specific sheet
```
/data sheet:Sales limit:5
```

### Search for "John"
```
/search query:John
```

### Get overview statistics
```
/stats overview
```

### Get statistics for the "Score" column
```
/stats column name:Score
```

### List all sheets
```
/sheets
```

## Project Structure

```
AIM-Discord-Bot/
├── commands/           # Command files
│   ├── data.js        # Display data command
│   ├── search.js      # Search command
│   ├── sheets.js      # List sheets command
│   └── stats.js       # Statistics command
├── data/              # Excel files directory
│   └── README.md      # Data folder instructions
├── config.js          # Bot configuration
├── deploy-commands.js # Deploy slash commands to Discord
├── excelReader.js     # Excel reading functionality
├── index.js           # Main bot file
├── .env.example       # Environment variables template
├── .gitignore         # Git ignore file
├── package.json       # Node.js dependencies
└── README.md          # This file
```

## Troubleshooting

### Bot doesn't respond to commands
- Make sure you've deployed the slash commands: `node deploy-commands.js`
- Check that the bot has proper permissions in your Discord server
- Ensure Message Content Intent is enabled in the Discord Developer Portal

### "Failed to load Excel file" error
- Verify the file path in `.env` is correct
- Make sure the Excel file exists in the specified location
- Check that the file is a valid .xlsx or .xls file

### Commands not showing up
- Run `node deploy-commands.js` to register commands
- Wait a few minutes for Discord to update
- Try kicking and re-inviting the bot

## Contributing

Feel free to submit issues and pull requests!

## License

ISC
