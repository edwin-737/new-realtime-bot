// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
    MessageFactory,
    TeamsActivityHandler,
    TeamsInfo,
} = require('botbuilder');
const {
    Graph
} = require('./graph/graph');
const state = {
    "LIST_TEAMS": 0,
    "LIST_CHANNELS": 1,
    "LIST_CONVERSATIONS": 2
}
class AnonymousBot extends TeamsActivityHandler {
    _graph = null;
    constructor() {
        super();
        this._graph = new Graph();
        this.onMessage(async (context, next) => {
            const userId = context.activity.from.id;
            const messageText = context.activity.text.trim().toLocaleLowerCase();
            /*for debugging, checking if graph api returns teams joined by user*/
            if (messageText.includes("listteams")) {
                await this._graph.getJoinedTeams(userId)
                    .then((res) => {
                        const data = res.data;
                        console.log(data);
                    });
                const activity = MessageFactory.text('logged teams, check log stream');
                await turnContext.sendActivity(activity);
                await next();
            }
            //for now test using a known teamsChannelId
            //Replace with teamsChannelId retrieved from some database, or Microsoft Graph
            const teamsChannelId = '19:68d2027a02114599be34f2bf95901699@thread.tacv2';
            const question = context.activity.text;

            //send the message to the desired channel
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

    async cardActivityAsync(context, isUpdate) {
        const cardActions = [
            {
                type: ActionTypes.MessageBack,
                title: 'My Teams',
                value: null,
                text: 'MyTeams'
            },
        ];

        if (isUpdate) {
            await this.sendUpdateCard(context, cardActions);
        } else {
            await this.sendWelcomeCard(context, cardActions);
        }
    }

    async sendUpdateCard(context, cardActions) {
        const data = context.activity.value;
        data.count += 1;
        cardActions.push({
            type: ActionTypes.MessageBack,
            title: 'Update Card',
            value: data,
            text: 'UpdateCardAction'
        });
        const card = CardFactory.heroCard(
            'Updated card',
            `Update count: ${data.count}`,
            null,
            cardActions
        );
        card.id = context.activity.replyToId;
        const message = MessageFactory.attachment(card);
        message.id = context.activity.replyToId;
        await context.updateActivity(message);
    }

    async sendWelcomeCard(context, cardActions) {
        const initialValue = {
            count: 0
        };
        cardActions.push({
            type: ActionTypes.MessageBack,
            title: 'Update Card',
            value: initialValue,
            text: 'UpdateCardAction'
        });
        const card = CardFactory.heroCard(
            'Welcome card',
            '',
            null,
            cardActions
        );
        await context.sendActivity(MessageFactory.attachment(card));
    }

}
module.exports.AnonymousBot = AnonymousBot;
