require('dotenv').config();
const { Client, GatewayIntentBits } = require('discord.js');

console.log('Testing Discord bot token...');
console.log('Token from env:', process.env.DISCORD_TOKEN ? `Present (${process.env.DISCORD_TOKEN.length} chars)` : 'MISSING');
console.log('Token starts with:', process.env.DISCORD_TOKEN ? process.env.DISCORD_TOKEN.substring(0, 10) + '...' : 'N/A');

const client = new Client({
    intents: [
        GatewayIntentBits.Guilds,
        GatewayIntentBits.GuildMessages,
        GatewayIntentBits.MessageContent,
    ],
});

client.once('ready', () => {
    console.log('✅ SUCCESS! Logged in as:', client.user.tag);
    console.log('Bot ID:', client.user.id);
    process.exit(0);
});

client.on('error', error => {
    console.error('❌ Client error:', error);
    process.exit(1);
});

console.log('Attempting login...');
client.login(process.env.DISCORD_TOKEN)
    .then(() => console.log('Login promise resolved'))
    .catch(error => {
        console.error('❌ Login failed:', error.message);
        console.error('Error code:', error.code);
        process.exit(1);
    });

setTimeout(() => {
    console.error('❌ TIMEOUT - Login took too long (30s)');
    process.exit(1);
}, 30000);
