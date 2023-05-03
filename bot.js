// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, MessageFactory, TeamsActivityHandler, CardFactory, ConsoleTranscriptLogger } = require('botbuilder');

class EchoBot extends TeamsActivityHandler {
    constructor() {
        super();
        this.onMessage(async (context, next) => {

            const replyText = `Echo: ${context.activity.text}`;
            console.log('onMessage called ' + replyText);
            await context.sendActivity(MessageFactory.text(replyText, replyText));
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
        // Listen to MembersAdded event, view https://docs.microsoft.com/en-us/microsoftteams/platform/resources/bot-v3/bots-notifications for more events
        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            const welcomeText = "welcome, the bot works on teams!"
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                if (membersAdded[cnt].id) {
                    await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
                    break;
                }
            }
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
