// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
    MessageFactory,
    TeamsActivityHandler,
    TeamsInfo,
    ActionTypes,
    CardFactory
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
    constructor() {
        super();
        this._graph = new Graph();
        this.onMessage(async (context, next) => {
            const userId = context.activity.from.aadObjectId;
            const messageText = context.activity.text.trim().toLocaleLowerCase();
            /*for debugging, checking if graph api returns teams joined by user*/
            if (messageText === "start") {
                var teams = [];
                this._graph.setUserId(userId);
                await this._graph.getJoinedTeams()
                    .then((retrievedTeams) => {
                        teams = retrievedTeams;
                    });
                await this.cardActivityAsync(context, teams);
            }
            /*else if (messageText === "channls") {
                var channels = [];
                await this._graph.getJoinedChannels()
            }*/
            else {
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
            }
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

    async cardActivityAsync(context, teamsList) {
        var cardActions = [];
        var cardActionTemplate = {
            type: ActionTypes.MessageBack,
            title: 'fill this',
            value: null,
            text: 'fillThis'
        }
        for (let idx = 0; idx < teamsList.length; idx++) {
            var newCardAction = cardActionTemplate;
            newCardAction.title = teamsList[idx].displayName;
            newCardAction.text = teamsList[idx].displayName;
        }
        await this.sendTeamCard(context, cardActions);
    }

    async sendTeamCard(context, cardActions) {
        const card = CardFactory.heroCard(
            'Choose team',
            `Choose a team`,
            null,
            cardActions
        );
        card.id = context.activity.replyToId;
        const message = MessageFactory.attachment(card);
        message.id = context.activity.replyToId;
        await context.updateActivity(message);
    }

}
module.exports.AnonymousBot = AnonymousBot;
