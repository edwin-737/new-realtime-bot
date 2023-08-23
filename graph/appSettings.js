
const settings = {
    //Replace with your bot resource's App ID
    'clientId': 'f3ca1bc9-f15d-4193-8e96-59229eded721',
    //Replace with an Azure Active Directory client secret
    'clientSecret': 'LbS8Q~V1qGRUCZ5pCDq-xKWEwM8PPDyTNyeBLbD8',
    //Replace with your own azure tenant ID
    'tenantId': '8e46ed74-357d-4717-8dff-512d29d26531',
    //Replace with your own azure tenant ID
    'authTenant': '8e46ed74-357d-4717-8dff-512d29d26531',
    'graphUserScopes': [
        'Team.ReadBasic.All',
        'User.Read.All'
    ]
};
module.exports = settings;