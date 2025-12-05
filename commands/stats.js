const { SlashCommandBuilder, EmbedBuilder } = require('discord.js');
const ExcelReader = require('../excelReader');
const config = require('../config');

module.exports = {
    data: new SlashCommandBuilder()
        .setName('stats')
        .setDescription('Get statistics from the Excel sheet')
        .addSubcommand(subcommand =>
            subcommand
                .setName('overview')
                .setDescription('Get overview statistics of the sheet')
                .addStringOption(option =>
                    option.setName('sheet')
                        .setDescription('The sheet name (optional)')
                        .setRequired(false)))
        .addSubcommand(subcommand =>
            subcommand
                .setName('column')
                .setDescription('Get statistics for a specific column')
                .addStringOption(option =>
                    option.setName('name')
                        .setDescription('The column name')
                        .setRequired(true))
                .addStringOption(option =>
                    option.setName('sheet')
                        .setDescription('The sheet name (optional)')
                        .setRequired(false))),

    async execute(interaction) {
        await interaction.deferReply();

        try {
            const reader = new ExcelReader(config.excelFilePath);

            const loadResult = await reader.loadFile();
            if (!loadResult) {
                return interaction.editReply('❌ Failed to load data source. Please check the configuration.');
            }

            const subcommand = interaction.options.getSubcommand();
            const sheetName = interaction.options.getString('sheet');
            const sheets = await reader.getSheetNames();
            const targetSheet = sheetName || sheets[0];

            if (subcommand === 'overview') {
                // Get overview statistics
                const rowCount = await reader.getRowCount(sheetName);
                const columns = await reader.getColumns(sheetName);

                const embed = new EmbedBuilder()
                    .setColor('#ff9900')
                    .setTitle(`📈 Overview Statistics - "${targetSheet}"`)
                    .addFields(
                        { name: '📊 Total Rows', value: `${rowCount}`, inline: true },
                        { name: '📋 Total Columns', value: `${columns.length}`, inline: true },
                        { name: '📄 Total Sheets', value: `${sheets.length}`, inline: true },
                        { name: '🏷️ Column Names', value: columns.join(', ') || 'None', inline: false },
                        { name: '📑 Available Sheets', value: sheets.join(', '), inline: false }
                    )
                    .setTimestamp();

                await interaction.editReply({ embeds: [embed] });

            } else if (subcommand === 'column') {
                // Get column statistics
                const columnName = interaction.options.getString('name');
                const stats = await reader.getColumnStats(columnName, sheetName);

                if (stats.error) {
                    return interaction.editReply(`❌ ${stats.error}`);
                }

                const embed = new EmbedBuilder()
                    .setColor('#9900ff')
                    .setTitle(`📊 Column Statistics - "${columnName}"`)
                    .setDescription(`From sheet: "${targetSheet}"`)
                    .addFields(
                        { name: '🔢 Count', value: `${stats.count}`, inline: true },
                        { name: '➕ Sum', value: `${stats.sum.toFixed(2)}`, inline: true },
                        { name: '📊 Average', value: `${stats.average.toFixed(2)}`, inline: true },
                        { name: '⬇️ Minimum', value: `${stats.min}`, inline: true },
                        { name: '⬆️ Maximum', value: `${stats.max}`, inline: true },
                        { name: '〰️ Median', value: `${stats.median}`, inline: true }
                    )
                    .setTimestamp();

                await interaction.editReply({ embeds: [embed] });
            }

        } catch (error) {
            console.error('[ERROR] Error in stats command:', error);
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

            if (args.length === 0 || args[0] === 'overview') {
                // Overview stats
                const sheetName = args[1] || null;
                const sheets = await reader.getSheetNames();
                const targetSheet = sheetName || sheets[0];
                const rowCount = await reader.getRowCount(sheetName);
                const columns = await reader.getColumns(sheetName);

                const embed = new EmbedBuilder()
                    .setColor('#ff9900')
                    .setTitle(`📈 Overview Statistics - "${targetSheet}"`)
                    .addFields(
                        { name: '📊 Total Rows', value: `${rowCount}`, inline: true },
                        { name: '📋 Total Columns', value: `${columns.length}`, inline: true },
                        { name: '📄 Total Sheets', value: `${sheets.length}`, inline: true },
                        { name: '🏷️ Column Names', value: columns.join(', ') || 'None', inline: false },
                        { name: '📑 Available Sheets', value: sheets.join(', '), inline: false }
                    )
                    .setTimestamp();

                await message.reply({ embeds: [embed] });

            } else if (args[0] === 'column' && args.length >= 2) {
                // Column stats
                const columnName = args[1];
                const sheetName = args[2] || null;
                const sheets = await reader.getSheetNames();
                const targetSheet = sheetName || sheets[0];
                const stats = await reader.getColumnStats(columnName, sheetName);

                if (stats.error) {
                    return message.reply(`❌ ${stats.error}`);
                }

                const embed = new EmbedBuilder()
                    .setColor('#9900ff')
                    .setTitle(`📊 Column Statistics - "${columnName}"`)
                    .setDescription(`From sheet: "${targetSheet}"`)
                    .addFields(
                        { name: '🔢 Count', value: `${stats.count}`, inline: true },
                        { name: '➕ Sum', value: `${stats.sum.toFixed(2)}`, inline: true },
                        { name: '📊 Average', value: `${stats.average.toFixed(2)}`, inline: true },
                        { name: '⬇️ Minimum', value: `${stats.min}`, inline: true },
                        { name: '⬆️ Maximum', value: `${stats.max}`, inline: true },
                        { name: '〰️ Median', value: `${stats.median}`, inline: true }
                    )
                    .setTimestamp();

                await message.reply({ embeds: [embed] });
            } else {
                message.reply('❌ Invalid usage. Examples:\n`!stats overview [sheet]`\n`!stats column <name> [sheet]`');
            }

        } catch (error) {
            console.error('[ERROR] Error in stats command:', error);
            message.reply(`❌ Error: ${error.message}`);
        }
    }
};
