const { SlashCommandBuilder, EmbedBuilder } = require('discord.js');
const ExcelReader = require('../excelReader');
const config = require('../config');

module.exports = {
    data: new SlashCommandBuilder()
        .setName('data')
        .setDescription('Display data from the Excel sheet')
        .addStringOption(option =>
            option.setName('sheet')
                .setDescription('The sheet name to read from (optional)')
                .setRequired(false))
        .addIntegerOption(option =>
            option.setName('limit')
                .setDescription('Number of rows to display (default: 10)')
                .setRequired(false)),

    async execute(interaction) {
        await interaction.deferReply();

        try {
            const reader = new ExcelReader(config.excelFilePath);

            const loadResult = await reader.loadFile();
            if (!loadResult) {
                return interaction.editReply('❌ Failed to load data source. Please check the configuration.');
            }

            const sheetName = interaction.options.getString('sheet');
            const limit = interaction.options.getInteger('limit') || 10;

            // Get available sheets
            const sheets = await reader.getSheetNames();

            // Determine which sheet to read
            const targetSheet = sheetName || sheets[0];

            if (sheetName && !sheets.includes(sheetName)) {
                return interaction.editReply(`❌ Sheet "${sheetName}" not found. Available sheets: ${sheets.join(', ')}`);
            }

            const data = await reader.getSheetData(targetSheet);

            if (data.length === 0) {
                return interaction.editReply('❌ No data found in the sheet.');
            }

            // Limit the number of rows
            const limitedData = data.slice(0, limit);

            // Create embed
            const embed = new EmbedBuilder()
                .setColor('#0099ff')
                .setTitle(`📊 Data from "${targetSheet}"`)
                .setDescription(`Showing ${limitedData.length} of ${data.length} rows`)
                .setTimestamp();

            // Add fields for each row
            limitedData.forEach((row, index) => {
                const rowText = Object.entries(row)
                    .map(([key, value]) => `**${key}:** ${value}`)
                    .join('\n');

                embed.addFields({
                    name: `Row ${index + 1}`,
                    value: rowText || 'Empty row',
                    inline: false
                });
            });

            // Add footer with available sheets
            embed.setFooter({ text: `Available sheets: ${sheets.join(', ')}` });

            await interaction.editReply({ embeds: [embed] });

        } catch (error) {
            console.error('[ERROR] Error in data command:', error);
            await interaction.editReply(`❌ Error: ${error.message}`);
        }
    },

    async executeMessage(message, args) {
        try {
            const reader = new ExcelReader(config.excelFilePath);

            const loadResult = await reader.loadFile();
            if (!loadResult) {
                return message.reply('❌ Failed to load data source.');
            }

            const limit = parseInt(args[0]) || 10;
            const sheetName = args[1] || null;

            const data = await reader.getSheetData(sheetName);

            if (data.length === 0) {
                return message.reply('❌ No data found in the sheet.');
            }

            const limitedData = data.slice(0, limit);
            const sheets = await reader.getSheetNames();
            const targetSheet = sheetName || sheets[0];

            const embed = new EmbedBuilder()
                .setColor('#0099ff')
                .setTitle(`📊 Data from "${targetSheet}"`)
                .setDescription(`Showing ${limitedData.length} of ${data.length} rows`)
                .setTimestamp();

            limitedData.forEach((row, index) => {
                const rowText = Object.entries(row)
                    .map(([key, value]) => `**${key}:** ${value}`)
                    .join('\n');

                embed.addFields({
                    name: `Row ${index + 1}`,
                    value: rowText || 'Empty row',
                    inline: false
                });
            });

            await message.reply({ embeds: [embed] });

        } catch (error) {
            console.error('[ERROR] Error in data command:', error);
            message.reply(`❌ Error: ${error.message}`);
        }
    }
};
