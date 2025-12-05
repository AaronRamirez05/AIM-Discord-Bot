const { SlashCommandBuilder, EmbedBuilder } = require('discord.js');
const ExcelReader = require('../excelReader');
const config = require('../config');

module.exports = {
    data: new SlashCommandBuilder()
        .setName('sheets')
        .setDescription('List all available sheets in the Excel file'),

    async execute(interaction) {
        await interaction.deferReply();

        try {
            const reader = new ExcelReader(config.excelFilePath);

            const loadResult = await reader.loadFile();
            if (!loadResult) {
                return interaction.editReply('❌ Failed to load data source. Please check the configuration.');
            }

            const sheets = await reader.getSheetNames();

            if (sheets.length === 0) {
                return interaction.editReply('❌ No sheets found in the data source.');
            }

            // Get row counts for all sheets
            const sheetFields = await Promise.all(
                sheets.map(async (sheet, index) => ({
                    name: `${index + 1}. ${sheet}`,
                    value: `Rows: ${await reader.getRowCount(sheet)}`,
                    inline: true
                }))
            );

            const embed = new EmbedBuilder()
                .setColor('#0099ff')
                .setTitle('📑 Available Sheets')
                .setDescription(`Found ${sheets.length} sheet(s):`)
                .addFields(sheetFields)
                .setTimestamp();

            await interaction.editReply({ embeds: [embed] });

        } catch (error) {
            console.error('[ERROR] Error in sheets command:', error);
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

            const sheets = await reader.getSheetNames();

            if (sheets.length === 0) {
                return message.reply('❌ No sheets found in the data source.');
            }

            // Get row counts for all sheets
            const sheetFields = await Promise.all(
                sheets.map(async (sheet, index) => ({
                    name: `${index + 1}. ${sheet}`,
                    value: `Rows: ${await reader.getRowCount(sheet)}`,
                    inline: true
                }))
            );

            const embed = new EmbedBuilder()
                .setColor('#0099ff')
                .setTitle('📑 Available Sheets')
                .setDescription(`Found ${sheets.length} sheet(s):`)
                .addFields(sheetFields)
                .setTimestamp();

            await message.reply({ embeds: [embed] });

        } catch (error) {
            console.error('[ERROR] Error in sheets command:', error);
            message.reply(`❌ Error: ${error.message}`);
        }
    }
};
