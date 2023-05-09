const { retrieveJoinedTeamsAsync, ensureGraphForAppOnlyAuth } = require('./graphHelper');

settings = require('./appSettings');

class Graph {
    constructor() {
        ensureGraphForAppOnlyAuth(settings);
    }
    async getJoinedTeams(id) {
        return retrieveJoinedTeamsAsync(id);
    }
}
module.exports.Graph = Graph;