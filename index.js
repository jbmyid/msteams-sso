var express = require('express');
var router = express.Router();

const {
    BotFrameworkAdapter,
    ConversationState,
    MemoryStorage,
    UserState
} = require('botbuilder');


const {
    MainDialog
} = require('./maindialog');

const {TeamsBot} = require("./teamsBot");
// Create adapter.
var adapter;
var botObj = {};

const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);
const userState = new UserState(memoryStorage);
const dialog = new MainDialog();
const teams_bot = new TeamsBot(conversationState, userState, dialog);


const errorConfig = {};

router.get('/', async function (req, res, next) {
    res.send('SSO running').status(200);
});

router.post('/', (req, res) => {
    try {
        adapter = new BotFrameworkAdapter({
            appId: req.query.app_id,
            appPassword: req.query.app_secret
        });
        console.log("MS TEAMS Request", req.body,req.query,req.headers);
        adapter.processActivity(req, res, async (context) => {
            // Route to main dialog.
            const botRunPromise = teams_bot.run(context);
            await (Promise.all([botRunPromise]));
            if(context.token){
                res.send({token: context.token, cRef: context.cRef}).status(200);
            }

            await (Promise.all([botRunPromise]));
        });

        adapter.onTurnError = async (context, error) => {
            // This check writes out errors to console log .vs. app insights.
            console.error(`\n [onTurnError] unhandled error: ${error}`);

            // Send a trace activity, which will be displayed in Bot Framework Emulator
            await context.sendTraceActivity(
                'OnTurnError Trace',
                `${error}`,
                'https://www.botframework.com/schemas/error',
                'TurnError'
            );


            await context.sendActivity('The bot encountered an error or bug.');
            await context.sendActivity('To continue to run this bot, please fix the bot source code.');
            // Clear out state
            await conversationState.delete(context);
        };
    } catch (e) {
        logger.error('Exception');
        logger.error(e);
        res.send(e).status(500);
    }
});

module.exports = {
    router,
    botObj
};