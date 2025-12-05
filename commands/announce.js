const { SlashCommandBuilder, EmbedBuilder } = require('discord.js');
const ExcelReader = require('../excelReader');
const config = require('../config');

module.exports = {
    data: new SlashCommandBuilder()
        .setName('announce-rankings')
        .setDescription('Post rankings announcement with all lifter stats'),

    async execute(interaction) {
        await interaction.deferReply();

        try {
            const reader = new ExcelReader(config.excelFilePath);

            const loadResult = await reader.loadFile();
            if (!loadResult) {
                return interaction.editReply('❌ Failed to load data source. Please check the configuration.');
            }

            const data = await reader.getSheetData();

            if (data.length === 0) {
                return interaction.editReply('❌ No lifter data found in the sheet.');
            }

            // Sort by Total Dots (highest first)
            const sortedData = [...data].sort((a, b) => {
                const dotsA = parseFloat(a['Total Dots']) || 0;
                const dotsB = parseFloat(b['Total Dots']) || 0;
                return dotsB - dotsA;
            });

            // Calculate overview statistics
            const totalLifters = sortedData.length;
            const totalDots = sortedData.reduce((sum, lifter) => sum + (parseFloat(lifter['Total Dots']) || 0), 0);
            const avgDots = totalDots / totalLifters;

            // Find highest individual lifts
            const maxSquat = Math.max(...sortedData.map(l => parseFloat(l['Squat - Single Max']) || 0));
            const maxBench = Math.max(...sortedData.map(l => parseFloat(l['Bench - Single Max']) || 0));
            const maxDeadlift = Math.max(...sortedData.map(l => parseFloat(l['Deadlift - Single Max']) || 0));

            // Create announcement message
            const announcementEmbed = new EmbedBuilder()
                .setColor('#FFD700')
                .setTitle('🏆 RANKINGS ARE COMPLETE! 🏆')
                .setDescription('@everyone The latest powerlifting rankings are now available!')
                .addFields(
                    { name: '👥 Total Participants', value: `${totalLifters}`, inline: true },
                    { name: '📊 Average Total Dots', value: `${avgDots.toFixed(2)}`, inline: true },
                    { name: '\u200B', value: '\u200B', inline: true },
                    { name: '🏋️ Highest Squat', value: `${maxSquat} lbs`, inline: true },
                    { name: '💪 Highest Bench', value: `${maxBench} lbs`, inline: true },
                    { name: '🔥 Highest Deadlift', value: `${maxDeadlift} lbs`, inline: true }
                )
                .setTimestamp();

            // Create rankings embed
            const rankingsEmbed = new EmbedBuilder()
                .setColor('#0099ff')
                .setTitle('📋 Final Rankings')
                .setDescription('Sorted by Total Dots Score')
                .setTimestamp();

            // Add each lifter to the rankings
            sortedData.forEach((lifter, index) => {
                const rank = index + 1;
                const medal = rank === 1 ? '🥇' : rank === 2 ? '🥈' : rank === 3 ? '🥉' : `${rank}.`;

                const name = lifter['Lifter'] || 'Unknown';
                const bodyweight = lifter['Bodyweight(Pounds)'] || 'N/A';
                const totalDots = lifter['Total Dots'] || '0';
                const squatMax = lifter['Squat - Single Max'] || '0';
                const benchMax = lifter['Bench - Single Max'] || '0';
                const deadliftMax = lifter['Deadlift - Single Max'] || '0';

                const lifterInfo = [
                    `**Total Dots:** ${totalDots}`,
                    `**Bodyweight:** ${bodyweight} lbs`,
                    `**S/B/D:** ${squatMax}/${benchMax}/${deadliftMax} lbs`
                ].join('\n');

                rankingsEmbed.addFields({
                    name: `${medal} ${name}`,
                    value: lifterInfo,
                    inline: false
                });
            });

            // Send both embeds
            await interaction.editReply({
                content: '@everyone',
                embeds: [announcementEmbed, rankingsEmbed]
            });

        } catch (error) {
            console.error('[ERROR] Error in announce-rankings command:', error);
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

            const data = await reader.getSheetData();

            if (data.length === 0) {
                return message.reply('❌ No lifter data found in the sheet.');
            }

            // Sort by Total Dots (highest first)
            const sortedData = [...data].sort((a, b) => {
                const dotsA = parseFloat(a['Total Dots']) || 0;
                const dotsB = parseFloat(b['Total Dots']) || 0;
                return dotsB - dotsA;
            });

            // Calculate overview statistics
            const totalLifters = sortedData.length;
            const totalDots = sortedData.reduce((sum, lifter) => sum + (parseFloat(lifter['Total Dots']) || 0), 0);
            const avgDots = totalDots / totalLifters;

            // Find highest individual lifts
            const maxSquat = Math.max(...sortedData.map(l => parseFloat(l['Squat - Single Max']) || 0));
            const maxBench = Math.max(...sortedData.map(l => parseFloat(l['Bench - Single Max']) || 0));
            const maxDeadlift = Math.max(...sortedData.map(l => parseFloat(l['Deadlift - Single Max']) || 0));

            // Create announcement message
            const announcementEmbed = new EmbedBuilder()
                .setColor('#FFD700')
                .setTitle('🏆 RANKINGS ARE COMPLETE! 🏆')
                .setDescription('@everyone The latest powerlifting rankings are now available!')
                .addFields(
                    { name: '👥 Total Participants', value: `${totalLifters}`, inline: true },
                    { name: '📊 Average Total Dots', value: `${avgDots.toFixed(2)}`, inline: true },
                    { name: '\u200B', value: '\u200B', inline: true },
                    { name: '🏋️ Highest Squat', value: `${maxSquat} lbs`, inline: true },
                    { name: '💪 Highest Bench', value: `${maxBench} lbs`, inline: true },
                    { name: '🔥 Highest Deadlift', value: `${maxDeadlift} lbs`, inline: true }
                )
                .setTimestamp();

            // Create rankings embed
            const rankingsEmbed = new EmbedBuilder()
                .setColor('#0099ff')
                .setTitle('📋 Final Rankings')
                .setDescription('Sorted by Total Dots Score')
                .setTimestamp();

            // Add each lifter to the rankings
            sortedData.forEach((lifter, index) => {
                const rank = index + 1;
                const medal = rank === 1 ? '🥇' : rank === 2 ? '🥈' : rank === 3 ? '🥉' : `${rank}.`;

                const name = lifter['Lifter'] || 'Unknown';
                const bodyweight = lifter['Bodyweight(Pounds)'] || 'N/A';
                const totalDots = lifter['Total Dots'] || '0';
                const squatMax = lifter['Squat - Single Max'] || '0';
                const benchMax = lifter['Bench - Single Max'] || '0';
                const deadliftMax = lifter['Deadlift - Single Max'] || '0';

                const lifterInfo = [
                    `**Total Dots:** ${totalDots}`,
                    `**Bodyweight:** ${bodyweight} lbs`,
                    `**S/B/D:** ${squatMax}/${benchMax}/${deadliftMax} lbs`
                ].join('\n');

                rankingsEmbed.addFields({
                    name: `${medal} ${name}`,
                    value: lifterInfo,
                    inline: false
                });
            });

            // Send both embeds
            await message.reply({
                content: '@everyone',
                embeds: [announcementEmbed, rankingsEmbed]
            });

        } catch (error) {
            console.error('[ERROR] Error in announce-rankings command:', error);
            message.reply(`❌ Error: ${error.message}`);
        }
    }
};
