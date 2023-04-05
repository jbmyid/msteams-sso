// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
  ConfirmPrompt,
  DialogSet,
  DialogTurnStatus,
  WaterfallDialog
} = require('botbuilder-dialogs');
const {
  LogoutDialog
} = require('./logoutdialog');

const CONFIRM_PROMPT = 'ConfirmPrompt';
const MAIN_DIALOG = 'MainDialog';
const MAIN_WATERFALL_DIALOG = 'MainWaterfallDialog';
const OAUTH_PROMPT = 'OAuthPrompt';
const {
  SsoOAuthPrompt
} = require('./ssooauthprompt');

const {
  TurnContext,
  TeamsInfo
} = require('botbuilder');


class MainDialog extends LogoutDialog {
  constructor() {
      super(MAIN_DIALOG, "Teams_SSO");
      this.addDialog(new SsoOAuthPrompt(OAUTH_PROMPT, {
          connectionName: "Teams_SSO",
          text: 'Please Sign In',
          title: 'Sign In',
          timeout: 300000
      }));
      this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));
      this.addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
          this.promptStep.bind(this),
          this.loginStep.bind(this)
      ]));

      this.initialDialogId = MAIN_WATERFALL_DIALOG;
  }

  /**
   * The run method handles the incoming activity (in the form of a DialogContext) and passes it through the dialog system.
   * If no dialog is active, it will start the default dialog.
   * @param {*} dialogContext
   */
  async run(context, accessor) {
      const dialogSet = new DialogSet(accessor);
      dialogSet.add(this);
      const dialogContext = await dialogSet.createContext(context);
      const results = await dialogContext.continueDialog();
      if (results.status === DialogTurnStatus.empty) {
          await dialogContext.beginDialog(this.id);
      }
  }

  async promptStep(stepContext) {
      try {
          let result = await stepContext.beginDialog(OAUTH_PROMPT);
          return result;
      } catch (err) {
          console.error(err);
      }
  }

  async loginStep(stepContext) {
      // Get the token from the previous step. Note that we could also have gotten the
      // token directly from the prompt itself. There is an example of this in the next method.
      const tokenResponse = stepContext.result;
      if (!tokenResponse || !tokenResponse.token) {
          console.log('Login Failed');
        //   await stepContext.context.sendActivity('Login was not successful. Sign out of the Teams application (by clicking on your profile picture), sign back in and try again.');
      } else {
          const cRef = TurnContext.getConversationReference(stepContext.context.activity);
          const user = await TeamsInfo.getMember(stepContext.context, encodeURI(stepContext.context.activity.from.id));
         
          console.log("config", config);
          stepContext.context.cRef= JSON.stringify(cRef),

          stepContext.context.token = tokenResponse.token;
          return await stepContext.endDialog();
      }
      return await stepContext.endDialog();
  }
}

module.exports.MainDialog = MainDialog;