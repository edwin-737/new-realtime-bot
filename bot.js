// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
    MessageFactory,
    TeamsActivityHandler,
    TeamsInfo,
    CardFactory,
    ActionTypes,
} = require('botbuilder');

class EchoBot extends TeamsActivityHandler {
    constructor() {
        super();
        this.onMessage(async (context, next) => {
            const replyText = `Echo: ${context.activity.text}`;
            console.log('onMessage called\n' + replyText);

            //for now test using a known teamsChannelId
            //Replace with teamsChannelId retrieved from some database 
            const teamsChannelId = '19:68d2027a02114599be34f2bf95901699@thread.tacv2';
            const question = context.activity.text;

            // //send the message to the desired channel
            const activity = MessageFactory.text(question);
            const [reference] = await TeamsInfo.sendMessageToTeamsChannel(context, activity, teamsChannelId, process.env.MicrosoftAppId);
            await context.adapter.continueConversationAsync(process.env.MicrosoftAppId, reference, async turnContext => {
                await turnContext.sendActivity(MessageFactory.text(question));
            });

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
        // Listen to MembersAdded event, view https://docs.microsoft.com/en-us/microsoftteams/platform/resources/bot-v3/bots-notifications for more events
        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            const welcomeText = "Welcome, this bot is for testing.";
            console.log('onMembers added called\n' + welcomeText);
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                if (membersAdded[cnt].id) {
                    await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
                    break;
                }
            }
            await next();
        });
    }
    // async cardActivityAsync(context, isUpdate) {
    //     const cardActions = [
    //         {
    //             type: ActionTypes.MessageBack,
    //             title: 'Message all members',
    //             value: null,
    //             text: 'MessageAllMembers'
    //         }
    //     ];

    //     if (isUpdate) {
    //         await this.sendUpdateCard(context, cardActions);
    //     } else {
    //         await this.sendWelcomeCard(context, cardActions);
    //     }
    // }
    // async sendWelcomeCard(context, cardActions) {
    //     const initialValue = {
    //         count: 0
    //     };
    //     cardActions.push({
    //         type: ActionTypes.MessageBack,
    //         title: 'Update Card',
    //         value: initialValue,
    //         text: 'UpdateCardAction'
    //     });
    //     const card = CardFactory.heroCard(
    //         'Welcome card',
    //         '',
    //         null,
    //         cardActions
    //     );
    //     await context.sendActivity(MessageFactory.attachment(card));
    // }

}
module.exports.EchoBot = EchoBot;
