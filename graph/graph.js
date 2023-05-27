const { retrieveJoinedTeamsAsync, ensureGraphForAppOnlyAuth, retrieveConversationsAsync, retrieveChannelsAsync } = require('./graphHelper');

settings = require('./appSettings');

class Graph {
    constructor() {
        this.user = null;
        this.chosenTeam = null;
        this.chosenChannel = null;
        this.chosenConversation = null;
        ensureGraphForAppOnlyAuth(settings);
    }
    getUser() {
        return this.user;
    }
    setUser(user) {
        this.user = user;
    }
    getChosenTeam() {
        return this.chosenTeam;
    }
    setChosenTeam(chosenTeam) {
        this.chosenTeam = chosenTeam;
    }
    getChosenChannel() {
        return this.chosenChannel;
    }
    setChosenChannel(chosenChannel) {
        this.chosenChannel = chosenChannel;
    }
    getChosenConversation() {
        return this.chosenConversation;
    }
    setChosenConversationId(chosenConversation) {
        this.chosenConversation = chosenConversation;
    }
    resetAllFields() {
        this.user = null;
        this.chosenTeam = null;
        this.chosenChannel = null;
        this.chosenConversation = null;
    }
    async getJoinedTeams() {
        return retrieveJoinedTeamsAsync(this.getUser().id);
    }
    async getJoinedChannels() {
        return retrieveChannelsAsync(this.getChosenTeam().id);
    }
    async getConversationsWithBot() {
        return retrieveConversationsAsync(this.getChosenChannel().id);
    }
}
module.exports.Graph = Graph;