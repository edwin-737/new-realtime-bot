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
const texts = require('./public/docs/texts.json');
class AnonymousBot extends TeamsActivityHandler {
    constructor() {
        super();
        this._graph = new Graph();
        this.onMessage(async (context, next) => {
            const messageText = context.activity.text.trim();
            /*give user info about the bot */
            if (messageText === '/help') {
                const helpText = texts.help;
                await context.sendActivity(MessageFactory.text(helpText, helpText));
            }
            /*let user choose which team to post question to */
            else if (messageText === '/start') {
                this._graph.resetAllFields();
                var teams = [];
                const id = context.activity.from.aadObjectId;
                const name = context.activity.from.name;
                // const id = 'a495e614-3794-4de3-847e-d2b6d4856c0b';
                // const name = 'conard samlu';
                const user = {
                    name: name,
                    id: id
                };
                this._graph.setUser(user);
                await this._graph.getJoinedTeams()
                    .then(async (retrievedTeams) => {
                        teams = retrievedTeams.value;
                    });
                console.log('teams retireved');
                console.log(teams);
                await this.cardActivityAsync(context, teams, 'teams');
            }
            /*let user choose which channel to post to */
            else if (messageText.startsWith('choose_channel/')) {
                var channels = [];
                const nameAndId = messageText.slice(messageText.indexOf('/') + 1, messageText.length);
                console.log(nameAndId);
                const team = this.createNameAndIdObject(nameAndId);
                this._graph.setChosenTeam(team);
                console.log('choosing channel for team:' + team.id + '/');
                await this._graph.getJoinedChannels()
                    .then(async (retrievedChannels) => {
                        channels = retrievedChannels.value;
                    });
                console.log('channels retrieved');
                console.log(channels);
                await this.cardActivityAsync(context, channels, 'channels');

            }
            else if (messageText.startsWith('send_message/')) {
                //for now test using a known teamsChannelId
                //Replace with teamsChannelId retrieved from some database, or Microsoft Graph
                const nameAndId = messageText.slice(messageText.indexOf('/') + 1, messageText.length);
                const channel = this.createNameAndIdObject(nameAndId);
                this._graph.setChosenChannel(channel);
                console.log(this._graph.getChosenChannel().id);
                const messageCard = CardFactory.heroCard(
                    'Send an <b>anonymous</b> message',
                    'Now you are ready to send your message in this chat. Your message will be routed to the channel' + ' <b>' + this._graph.getChosenChannel().name + '</b> in the team ' + '<b>' + this._graph.getChosenTeam().name + '</b>.If you want to change the team or channel, send <b>/start</b> again to restart the selection process.',
                    null,
                    null,
                )
                const message = MessageFactory.attachment(messageCard);
                await context.sendActivity(message);
            }
            else {
                if (!this._graph.getChosenChannel()) {
                    await context.sendActivity(MessageFactory.text('Select a team and channel first. Send a **/start** command to begin choosing.'));
                }
                else {
                    const question = context.activity.text;
                    //send the message to the desired channel
                    const activity = MessageFactory.text(question);
                    const teamsChannelId = this._graph.getChosenChannel().id;
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
            const welcomeText = texts.welcome;
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                if (membersAdded[cnt].id) {
                    await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
                    break;
                }
            }
            await next();
        });
    }
    createNameAndIdObject(nameAndId) {
        const id = nameAndId.slice(nameAndId, nameAndId.indexOf('/'));
        const name = nameAndId.slice(nameAndId.indexOf('/') + 1, nameAndId.length);
        const nameAndIdObj = {
            name: name,
            id: id
        }
        return nameAndIdObj;
    }
    async cardActivityAsync(context, genericList, typeOfCard) {
        var cardActions = new Array(genericList.length);
        // console.log(genericList);
        var trailingText = '';
        if (typeOfCard === 'channels')
            trailingText = 'send_message/';
        else if (typeOfCard == 'teams')
            trailingText = 'choose_channel/'
        //Populate cardActions using the list of teams passed in
        for (var idx = 0; idx < genericList.length; idx++) {
            var newCardAction = {
                type: ActionTypes.MessageBack,
                title: 'fill this',
                value: null,
                text: 'fillThis'
            };
            newCardAction.title = genericList[idx].displayName;
            newCardAction.text = trailingText + genericList[idx].id + '/' + genericList[idx].displayName;
            cardActions[idx] = newCardAction;
        }
        if (typeOfCard === 'channels')
            await this.sendChannelCard(context, cardActions);
        else if (typeOfCard == 'teams')
            await this.sendTeamCard(context, cardActions);

    }

    async sendTeamCard(context, cardActions) {
        const card = CardFactory.heroCard(
            'Choose a team',
            'Choose the team to send your message to',
            null,
            cardActions
        );
        card.id = context.activity.id;
        const message = MessageFactory.attachment(card);
        // message.id = context.activity.id;
        await context.sendActivity(message);
    }
    async sendChannelCard(context, cardActions) {
        const card = CardFactory.heroCard(
            'Choose a channel',
            'Choose the channel from the team ' + '<b>' + this._graph.getChosenTeam().name + '</b>',
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
