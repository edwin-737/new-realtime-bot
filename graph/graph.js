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
    async getJoinedTeams() {
        return retrieveJoinedTeamsAsync(userId);
    }
    async getJoinedChannels(chosenTeamId) {
        return retrieveChannelsAsync(chosenTeamId);
    }
    async getConversationsWithBot(chosenChannelId) {
        return retrieveConversationsAsync(chosenChannelId);
    }
}
module.exports.Graph = Graph;