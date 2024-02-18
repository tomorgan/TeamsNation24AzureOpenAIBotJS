const axios = require("axios");

const { TeamsActivityHandler, TurnContext } = require("botbuilder");

const URL =
  "https://{OPENAIINSTANCE}.openai.azure.com/openai/deployments/{DEPLOYMENTNAME}/extensions/chat/completions?api-version=2023-06-01-preview";

const AnswerQuestion = async (question) => {
  const postData = {
    dataSources: [
      {
        type: "AzureCognitiveSearch",
        parameters: {
          endpoint: "{Azure Search Endpoint URL}",
          key: "{Azure Search Endpoint Key}",
          indexName: "{Azure Search Index Name}",
          semanticConfiguration: "",
          queryType: "simple",
          fieldsMapping: {
            contentFieldsSeparator: "\n",
            contentFields: [
              "List",
              "Of",
              "Content",
              "Fields",
              "From",
              "Azure",
              "Search",
              "Service",
            ],
            filepathField: "id", //change these if needed to match the values in Azure Search Service
            titleField: "title", //change these if needed to match the values in Azure Search Service
            urlField: "url", //change these if needed to match the values in Azure Search Service
          },
          inScope: true,
          roleInformation: "{Add Prompt}",
        },
      },
    ],
    messages: [
      {
        role: "user",
        content: question,
      },
    ],
    deployment: "{Deployment Name}",
    temperature: 0,
    top_p: 1,
    max_tokens: 800,
    stop: null,
    stream: false,
  };

  const response = await axios.post(URL, postData, {
    headers: {
      "api-key": "{Azure OpenAI API Key}",
      "Content-Type": "application/json",
    },
  });

  return response.data.choices[0].messages.find((m) => m.role === "assistant")
    .content;
};

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
      );
      const txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      const answerTxt = await AnswerQuestion(txt);
      //await context.sendActivity(`Echo: ${txt}`);
      await context.sendActivity(answerTxt);

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    // Listen to MembersAdded event, view https://docs.microsoft.com/en-us/microsoftteams/platform/resources/bot-v3/bots-notifications for more events
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          await context.sendActivity(
            `Hi there! I'm a Teams bot that will echo what you said to me.`
          );
          break;
        }
      }
      await next();
    });
  }
}

module.exports.TeamsBot = TeamsBot;
