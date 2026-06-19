const fs = require('fs');
const path = require('path');

const configPath = path.join(__dirname, '../data/config.json');

function loadConfig() {
    try {
        return JSON.parse(fs.readFileSync(configPath, 'utf8'));
    } catch (err) {
        console.error("❌ Failed to load config.json", err);
        return {};
    }
}

function saveConfig(newConfig) {
    fs.writeFileSync(configPath, JSON.stringify(newConfig, null, 2), 'utf8');
}

module.exports = { loadConfig, saveConfig };