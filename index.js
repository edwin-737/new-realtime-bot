// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
const path = require('path');
const dotenv = require('dotenv');
// Import required bot configuration.
const ENV_FILE = path.join(__dirname, '.env');
dotenv.config({ path: ENV_FILE });

const restify = require('restify');

//Import graph
const { Graph } = require('./graph/graph')
require('./graph/graphHelper');
const myGraph = new Graph();
// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
const {
    CloudAdapter,
    ConfigurationBotFrameworkAuthentication
} = require('botbuilder');

// This bot's main dialog.
const { AnonymousBot } = require('./bot');
const { retrieveJoinedTeamsAsync, retrieveChannelsAsync, retrieveConversationsAsync } = require('./graph/graphHelper');

// Create HTTP server
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`\n${server.name} listening to ${server.url}`);
    console.log('\nGet Bot Framework Emulator: https://aka.ms/botframework-emulator');
    console.log('\nTo talk to your bot, open the emulator select "Open Bot"');
});

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(process.env);

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about how bots work.
const adapter = new CloudAdapter(botFrameworkAuthentication);

// Catch-all for errors.
const onTurnErrorHandler = async (context, error) => {
    // This check writes out errors to console log .vs. app insights.
    // NOTE: In production environment, you should consider logging this to Azure
    //       application insights. See https://aka.ms/bottelemetry for telemetry
    //       configuration instructions.
    console.error(`\n [onTurnError] unhandled error: ${error}`);

    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${error}`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    // Send a message to the user
    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create the main dialog.
const myBot = new AnonymousBot();

// Listen for incoming requests.
server.post('/api/messages', async (req, res) => {
    console.log(req.body);
    // Route received a request to adapter for processing
    await adapter.process(req, res, async (context) => await myBot.run(context));
});
//for testing graph api
server.post('/api/graph/teams', async (req, res) => {
    await myGraph.getJoinedTeams(req.body.id)
        .then(teams => {
            console.log('retrieved users joined teams');
            console.log(teams);
            res.send(teams);
        });
});
server.post('/api/graph/channels', async (req, res) => {
    await myGraph.getJoinedChannels(req.body.id)
        .then(channels => {
            console.log('retireved channels in team');
            console.log(channels);
            res.send(channels);
        });
});
server.post('/api/graph/conversations', async (req, res) => {
    await myGraph.getConversationsWithBot(req.body.id)
        .then(conversations => {
            console.log('retireved conversations in channel');
            console.log(conversations);
            res.send(conversations);
        });
});