const express = require('express');
const { BotFrameworkAdapter } = require('botbuilder');
const bodyParser = require('body-parser');

const app = express();
app.use(bodyParser.json());

const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
});

const port = process.env.PORT || 3978;

// Handle incoming messages
app.post('/api/messages', (req, res) => {
    adapter.processActivity(req, res, async (context) => {
        if (context.activity.type === 'message') {
            await context.sendActivity(`You said: ${context.activity.text}`);
        }
    });
});

// Start the server
app.listen(port, () => {
    console.log(`Server is listening on port ${port}`);
});
