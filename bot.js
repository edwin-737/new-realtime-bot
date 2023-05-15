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
                this._graph.setUserId('a495e614-3794-4de3-847e-d2b6d4856c0b');
                await this._graph.getJoinedTeams()
                    .then(async (retrievedTeams) => {
                        teams = retrievedTeams.value;
                    });
                console.log('teams retireved');
                console.log(teams);
                await this.cardActivityAsync(context, teams);
            }
            else if (messageText === "choose_channel") {
                var channels = [];
                console.log('choosing team...');
                const message = MessageFactory.text("choosing channel");
                await this.sendActivity(message);

                // await this._graph.getJoinedChannels()
            }
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
        console.log(teamsList);
        //Populate cardActions using the list of teams passed in
        for (let idx = 0; idx < teamsList.length; idx++) {
            var newCardAction = cardActionTemplate;
            console.log(idx + ' ' + teamsList[idx]);
            newCardAction.title = teamsList[idx].displayName;
            newCardAction.text = teamsList[idx].displayName;
            cardActions.push(newCardAction);
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
        card.id = context.activity.id;
        const message = MessageFactory.attachment(card);
        message.text = "choose_channel"
        message.id = context.activity.id;
        await context.sendActivity(message);
    }

}
module.exports.AnonymousBot = AnonymousBot;
