const { SlashCommandBuilder, EmbedBuilder } = require('discord.js');
const ExcelReader = require('../excelReader');
const config = require('../config');

module.exports = {
    data: new SlashCommandBuilder()
        .setName('search')
        .setDescription('Search for data in the Excel sheet')
        .addStringOption(option =>
            option.setName('query')
                .setDescription('The search term')
                .setRequired(true))
        .addStringOption(option =>
            option.setName('sheet')
                .setDescription('The sheet name to search in (optional)')
                .setRequired(false)),

    async execute(interaction) {
        await interaction.deferReply();

        try {
            const reader = new ExcelReader(config.excelFilePath);

            const loadResult = await reader.loadFile();
            if (!loadResult) {
                return interaction.editReply('❌ Failed to load data source. Please check the configuration.');
            }

            const query = interaction.options.getString('query');
            const sheetName = interaction.options.getString('sheet');

            const results = await reader.search(query, sheetName);

            if (results.length === 0) {
                return interaction.editReply(`❌ No results found for "${query}"`);
            }

            const sheets = await reader.getSheetNames();
            const targetSheet = sheetName || sheets[0];

            // Limit results to 10
            const limitedResults = results.slice(0, 10);

            const embed = new EmbedBuilder()
                .setColor('#00ff00')
                .setTitle(`🔍 Search Results for "${query}"`)
                .setDescription(`Found ${results.length} result(s) in "${targetSheet}" (showing first ${limitedResults.length})`)
                .setTimestamp();

            limitedResults.forEach((row, index) => {
                const rowText = Object.entries(row)
                    .map(([key, value]) => `**${key}:** ${value}`)
                    .join('\n');

                embed.addFields({
                    name: `Result ${index + 1}`,
                    value: rowText || 'Empty',
                    inline: false
                });
            });

            await interaction.editReply({ embeds: [embed] });

        } catch (error) {
            console.error('[ERROR] Error in search command:', error);
            await interaction.editReply(`❌ Error: ${error.message}`);
        }
    },

    async executeMessage(message, args) {
        try {
            if (args.length === 0) {
                return message.reply('❌ Please provide a search term. Usage: `!search <term> [sheet]`');
            }

            const reader = new ExcelReader(config.excelFilePath);

            const loadResult = await reader.loadFile();
            if (!loadResult) {
                return message.reply('❌ Failed to load data source.');
            }

            const query = args[0];
            const sheetName = args[1] || null;

            const results = await reader.search(query, sheetName);

            if (results.length === 0) {
                return message.reply(`❌ No results found for "${query}"`);
            }

            const sheets = await reader.getSheetNames();
            const targetSheet = sheetName || sheets[0];
            const limitedResults = results.slice(0, 10);

            const embed = new EmbedBuilder()
                .setColor('#00ff00')
                .setTitle(`🔍 Search Results for "${query}"`)
                .setDescription(`Found ${results.length} result(s) in "${targetSheet}" (showing first ${limitedResults.length})`)
                .setTimestamp();

            limitedResults.forEach((row, index) => {
                const rowText = Object.entries(row)
                    .map(([key, value]) => `**${key}:** ${value}`)
                    .join('\n');

                embed.addFields({
                    name: `Result ${index + 1}`,
                    value: rowText || 'Empty',
                    inline: false
                });
            });

            await message.reply({ embeds: [embed] });

        } catch (error) {
            console.error('[ERROR] Error in search command:', error);
            message.reply(`❌ Error: ${error.message}`);
        }
    }
};
