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
            // const userId = "a495e614-3794-4de3-847e-d2b6d4856c0b";
            const messageText = context.activity.text.trim();
            /*let user choose which team to post question to */
            if (messageText === "start") {
                var teams = [];
                this._graph.setUserId(userId);
                // this._graph.setUserId("a495e614-3794-4de3-847e-d2b6d4856c0b");
                await this._graph.getJoinedTeams(userId)
                    .then(async (retrievedTeams) => {
                        teams = retrievedTeams.value;
                    });
                console.log('teams retireved');
                console.log(teams);
                await this.cardActivityAsync(context, teams, "teams");
            }
            /*let user choose which channel to post to */
            else if (messageText.startsWith("choose_channel/")) {
                var channels = [];
                var teamId = messageText.replace("choose_channel/", "")
                this._graph.setChosenTeamId(teamId);
                console.log('choosing channel for team:' + teamId + "/");
                await this._graph.getJoinedChannels(teamId)
                    .then(async (retrievedChannels) => {
                        channels = retrievedChannels.value;
                    });
                console.log('channels retrieved');
                console.log(channels);
                await this.cardActivityAsync(context, channels, "channels");

            }
            else if (messageText.startsWith("send_message/")) {
                //for now test using a known teamsChannelId
                //Replace with teamsChannelId retrieved from some database, or Microsoft Graph
                const teamsChannelId = messageText.replace("send_message/", "");
                this._graph.setChosenChannelId(teamsChannelId);

                // const question = context.activity.text;

                await context.sendActivity(MessageFactory.text('Send your message now'));

                // //send the message to the desired channel
                // const activity = MessageFactory.text(question);
                // const [reference] = await TeamsInfo.sendMessageToTeamsChannel(context, activity, teamsChannelId, process.env.MicrosoftAppId);
                // await context.adapter.continueConversationAsync(process.env.MicrosoftAppId, reference, async turnContext => {
                //     await turnContext.sendActivity(MessageFactory.text(question));
                // });
            }
            else {
                if (this._graph.getChosenChannelId() === '') {
                    await context.sendActivity(MessageFactory.text('Select a team and channel first. Send a "start" command.'));
                }
                else {
                    const question = context.activity.text;
                    //send the message to the desired channel
                    const activity = MessageFactory.text(question);
                    const teamsChannelId = this._graph.getChosenChannelId();
                    console.log(teamsChannelId);
                    const [reference] = await TeamsInfo.sendMessageToTeamsChannel(context, activity, teamsChannelId, process.env.MicrosoftAppId);
                    await context.adapter.continueConversationAsync(process.env.MicrosoftAppId, reference, async turnContext => {
                        await turnContext.sendActivity(MessageFactory.text(question));
                    });
                }
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

    async cardActivityAsync(context, genericList, typeOfCard) {
        var cardActions = new Array(genericList.length);
        var cardActionTemplate =
            console.log(genericList);
        var trailingText = "";
        if (typeOfCard === "channels")
            trailingText = 'send_message/';
        else if (typeOfCard == "teams")
            trailingText = "choose_channel/"
        //Populate cardActions using the list of teams passed in
        for (var idx = 0; idx < genericList.length; idx++) {
            var newCardAction = {
                type: ActionTypes.MessageBack,
                title: 'fill this',
                value: null,
                text: 'fillThis'
            };
            newCardAction.title = genericList[idx].displayName;
            newCardAction.text = trailingText + genericList[idx].id;
            cardActions[idx] = newCardAction;
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
        // message.id = context.activity.id;
        await context.sendActivity(message);
    }

}
module.exports.AnonymousBot = AnonymousBot;
