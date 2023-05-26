const { retrieveJoinedTeamsAsync, ensureGraphForAppOnlyAuth, retrieveConversationsAsync, retrieveChannelsAsync } = require('./graphHelper');

settings = require('./appSettings');

class Graph {
    constructor() {
        this.userId = '';
        this.chosenTeamId = '';
        this.chosenChannelId = '';
        this.chosenConversationId = '';
        ensureGraphForAppOnlyAuth(settings);
    }
    getUserId() {
        return this.userId;
    }
    setUserId(userId) {
        this.userId = userId;
    }
    getChosenTeamId() {
        return this.chosenTeamId;
    }
    setChosenTeamId(chosenTeamId) {
        this.chosenTeamId = chosenTeamId;
    }
    getChosenChannelId() {
        return this.chosenChannelId;
    }
    setChosenChannelId(chosenChannelId) {
        this.chosenChannelId = chosenChannelId;
    }
    getChosenConversationId() {
        return this.chosenConversationId;
    }
    setChosenConversationId(chosenConversationId) {
        this.chosenConversationId = chosenConversationId;
    }
    async getJoinedTeams(user) {
        return retrieveJoinedTeamsAsync(user);
    }
    async getJoinedChannels(chosenTeam) {
        return retrieveChannelsAsync(chosenTeam);
    }
    async getConversationsWithBot(chosenChannel) {
        return retrieveConversationsAsync(chosenChannel);
    }
}
module.exports.Graph = Graph;