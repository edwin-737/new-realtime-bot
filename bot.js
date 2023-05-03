// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, MessageFactory, TeamsActivityHandler, CardFactory, ConsoleTranscriptLogger } = require('botbuilder');

class EchoBot extends TeamsActivityHandler {
    constructor() {
        super();
        this.onTeamsChannelCreatedEvent(async (channelInfo, teamInfo, turnContext, next) => {
            const card = CardFactory.heroCard('Channel Created', `${channelInfo.name} is the Channel created`);
            console.log('Channel created,' + channelInfo.name + 'is the Channel created');
            const message = MessageFactory.attachment(card);
            // Sends a message activity to the sender of the incoming activity.
            await turnContext.sendActivity(message);
            await next();
        });
        this.onMessage(async (context, next) => {

            const replyText = `Echo: ${context.activity.text}`;
            console.log('onMessage called ' + replyText);
            await context.sendActivity(MessageFactory.text(replyText, replyText));
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }
}

// class EchoBot extends ActivityHandler {
//     constructor() {
//         console.log('constructor called');
//         super();
//         // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
//         this.onMessage(async (context, next) => {

//             const replyText = `Echo: ${context.activity.text}`;
//             console.log('onMessage called ' + replyText);
//             await context.sendActivity(MessageFactory.text(replyText, replyText));
//             // By calling next() you ensure that the next BotHandler is run.
//             await next();
//         });

//         this.onMembersAdded(async (context, next) => {
//             const membersAdded = context.activity.membersAdded;
//             const welcomeText = 'Hello and welcome! New deployment';
//             console.log('onMembersAdded called');
//             for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
//                 if (membersAdded[cnt].id !== context.activity.recipient.id) {
//                     await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
//                 }
//             }
//             // By calling next() you ensure that the next BotHandler is run.
//             await next();
//         });
//     }
// }

module.exports.EchoBot = EchoBot;
