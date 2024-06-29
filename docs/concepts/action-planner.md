# Action Planner

The Action Planner is one the main components of [Teams AI Library](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/teams%20conversational%20ai/teams-conversation-ai-overview#action-planner) that calls your Large Language Model (LLM) to enhance and customize your model. Based on the LLM response, it generates the JSON formated command/response structured plan that is used to orchestrate execution of various custom parameterized actions that can be run sequentially or in parallel.

The Action Planner has a built-in prompt management system that supports creating prompt templates as folders in the file system. A prompt template is the dynamic prompt text along with all the configurations for completion with the LLM model. Dynamic prompts support template variables and functions.

Teams Copilot Starter comes with the following prompt templates:

- **plan**: uses a dynamic prompt and configuration parameters to control behavior of Teams AI Library Action Planner. This prompt template contains the configuration of all supported `actions` by Teams Copilot Starter. More about supported `actions` later in this document.
  
- **chatGPT**: uses a user defined dynamic prompt to support a freely flowing conversation with the user via use of LLM chat completion model. It has the ability to "remember" (maintain) the conversation history up to the configurable number of turns.
  
- **findEntity**: uses a user defined prompt template that instructs LLM to look up for a publicly known company entity to be returned in the specified JSON format:

  ```json
  {"company":"<company>", "info":"<info>", "ticker":"<ticker>", "address":"<address>", "city":"<city>", "country":"<country>", "website":"<website>", "revenue":"<revenue>", "industry":"<industry>", "employees":"<employees>"}
  ```

- **listEntities**: uses a user defined prompt template that instructs LLM to return the list of company names and their basic public information matching the human's text entry. The result is returned in the following JSON format:

  ```json
  {"entities":[{"company":"<company1>", "ticker":"<ticker1>", "location":"<location1>", "website":"<website1>"},{"company":"<company2>", "ticker":"<ticker2>","location":"<location2>", "website":"<website2>"},...]}
  ```

- **matchingEntities**: uses a user defined prompt template that is almost identical to the `listEntities` prompt template. The only difference is that it instructs LLM to return the list of companies whose names start with the exact letters from the human's text entry. For example, this prompt is used to look up company names that start with the user's entered letters in the Bot's message extension search box.

- **generatePrompts**: uses a user defined prompt template that instructs LLM to return up to 4 questions related to the company name found in the human's entered text. It returns the questions in the following format: `["question","question","question","question"]`

- **questionDocument**: uses a user defined prompt template that is designed to inject contextual data sources to augment the prompt further and facilitate better responses. The external data sources can be either in text or PDF file format that are typically uploaded from the user's local machine. Alternatively, the previously uploaded documents can be found and used from the Teams documents library. The data source name used by this prompt template is defined in the `augmentation` section of the configuration file `config.json` (as shown below), and its name must match the name of the corresponding environment variable `DOCUMENTDATA_SOURCE_NAME` defined in `.env.[local|dev|prod]` when deploying from the local environment or in the Bot Service Configuration section when your bot is deployed to Azure. See the example of how the data source named `documentdata` is defined in both, `config.json` and `.env.[local|dev|prod]`:

  `config.json`:

  ```json

    ...
    "augmentation": {
      "augmentation_type": "sequence",
      "data_sources": {
        "documentdata": 1200
      }
    }
  ```

  `.env.dev`:

  ```bash

  ...
  DOCUMENTDATA_SOURCE_NAME=documentdata
  ...
  ```

- **questionWeb**: uses a user defined prompt template designed to inject contextual web content to augment prompt and facilitate the responses that use web content data. The web content must come from a public website, containing relatively small amount of content available on that webpage. The maximum number of tokens allowed to be retrieved from the web content is configured in the `config.json` file as it's shown in the sample below. The web content must be mostly compiled from the text content, as images won't be taken into the prompt's retrieval augmentation generation (RAG). The data source name used by this prompt template is defined in the `augmentation` section of the configuration file `config.json` (as shown below), and its name must match the name of the corresponding environment variable `WEBDATA_SOURCE_NAME` defined in `.env.[local|dev|prod]` when deploying from the local environment or in the Bot Service Configuration section when your bot is deployed to Azure. See the example of how the data source named `webdata` is defined in both, `config.json` and `.env.[local|dev|prod]`:

  `config.json`:

  ```json

    ...
    "augmentation": {
      "augmentation_type": "sequence",
      "data_sources": {
        "webdata": 1200
      }
    }
  ```

  `.env.dev`:

  ```bash

  ...
  WEBDATA_SOURCE_NAME=webdata
  ...
  ```

## Actions

The action planner receives the user's ask and returns a plan on how to accomplish the request. The user's ask triggers the Teams Copilot Starter `plan` prompt template, which generates the action plan in the following JSON format:

```json
{"type": "plan", "commands": [{"type": "SAY", "response": "<message to user>"}, {"type": "DO", "action": "<action name>", "parameters": {"entity": "<action parameters>"}]}
```

The Action Planner does this by using Teams AI to mix and match atomic functions (called `actions`) registered in the Teams Copilot Starter in order to combine them into a series of steps that complete a goal. These `actions` are first described in the plan's `actions.json` metadata file and then registered in `TeamsAI` class as Teams AI handlers.

This is a powerful concept because it allows you to create actions that can be used in ways that you as a developer may not have thought of.

For instance, If you have a task with `getSemanticInfo` & `getCompanyDetails` actions, the planner could combine them to create workflows like "Tell me about Microsoft and provide their financial outlook" without you explicitly having to write code for those scenarios.

### Augmentations

Augmentations virtually eliminate the need for prompt engineering. Prompts can be configured to use named augmentation types which will instruct the action planner how to execute the given action. Augmentation types let the developer specify whether they want to support multi-step plans (`sequence`), run actions in parallel (`parallelActions`) or create an AutoGPT style agent (`monologue`).

Only `sequence` and `monologue` augmentation types can be defined in the prompt template `config.json` file.

```json
"augmentation": {
    "augmentation_type": "sequence"
  }
```

The `parallelActions` augmentation type is configured directly for each individual action in the `canRunWith` property. Refer to the [Parallel Actions](#parallel-actions) section for more details on how to configure this augmentation type.

#### Sequence Augmentation

This augmentation allows the model to return a sequence of actions to perform. It does this by appending instructions to the prompt text during runtime. These instructions guide the model to generate a plan object that uses actions defined in the `actions.json` file from the prompt template folder.

#### Monologue Augmentation

This augmentation adds support for an inner monologue to the prompt. The monologue helps the LLM perform chain-of-thought reasoning across multiple turns of conversation. It does this by appending instructions to the prompt text during runtime. It tells the model to explicitly show it's thought, reasoning and plan in response to the user's message, then predict the next action to execute. If looping is configured, then the predicted action can guide the model to predict the next action by returning the instruction as a string in the action handler callback. The loop will terminate as soon as the model predicts a `SAY` action, which sends the response back to the user.

#### Parallel Actions

This augmentation is unique compared to the previously described augmentation modes in the way that while both `sequence` and `monologue` augmentations are applied to the entire plan, the `parallelActions` is only retained within the action where it was declared. To declare that particular action can run in parallel with another action, you need to add the `canRunWith` property to the specific `action`. This new property takes an array of the action names that tell LLM that these actions can run in parallel with this parent action. For example, the out of box Teams Copilot Starter comes with these actions:

```json
  {
    "name": "getSemanticInfo",
    "description": "Retrieves the GPT response on the inquired company",
    "canRunWith": ["getCompanyDetails"],
    "parameters": {
      "type": "object",
      "properties": {
        "entity": {
          "type": "string",
          "description": "The company name for which the GPT response is being retrieved"
        }
      },
      "required": ["entity"]
    }
  },
  {
    "name": "getCompanyDetails",
    "description": "Gets the details about the specified company",
    "canRunWith": ["getSemanticInfo"],
    "parameters": {
      "type": "object",
      "properties": {
        "entity": {
          "type": "string",
          "description": "The company name for which the detailed information is being retrieved"
        }
      },
      "required": ["entity"]
    }
  },

```

The `canRunWith` property is declared in both actions. That means that when the plan creates to run both actions, most likely the `getSemanticInfo` will be the first to run. The plan will be created with only `getSemanticInfo` action in the actions list of the action plan, while `getCompanyDetails` action will be moved to be the _child_ of the `getSemanticInfo` action marked as a parallel action. You can see the resulted action plan in the example diagram below.

#### Action Commands and Entities

A plan consists of two types of commands and their entities:

- **SAY**: Sends a message to the user.
  - `response`: The string message to send.
- **DO**: AI system will execute a specific `action`, passing in the generated parameters.
  - `action`: A lambda function registered to the AI system
  - `parameters`: A dictionary passed to the action.
  - `parallelActions`: An array of `DO` actions, each can also be accompanied by `parameters` for that action.

The JSON object string is returned by the LLM and deserialized into an object.
Here's an example of a plan for the following user's prompt:

User: "`tell me about Walmart and then provide their financial outlook`"

![action plan](../images/screenshots/action-plan.png)

This plan is executed in parallel order as the `parallelActions` property of type `DO` suggests.

- The bot will send the first `SAY` message back to the user.
- [Parallel] The `getSemanticInfo` action will be executed using the chat GPT prompt template
- [Parallel] The `getCompanyDetails` action will be executed, which will return the company's financial details: company's annual revenue, # of employees, the latest info, ESG scores, suggested prompts and the list of other competitive companies. 
- The `response` message for the second type `SAY` will be sent to the user.

>Note how the parameters of the second action matches the parameter `entity` of the first action. The Teams Copilot Starter middleware class `ActionPlannerMiddleware` chains together the results of the currently completed action with the parameters of the next action to be executed per the action plan.
Let's assume that you uploaded the document where a company is being mentioned, but you don't know the company name mentioned in that document. Now, you want the Teams Copilot Starter bot to give you the financial details about that company. You could ask the bot: `extract the company name from this document and then provide me with their financial details`.

The planner is an extensible part of the Teams AI Library. This means that a custom planner can be created for your specific needs. Out of the box, the Teams Copilot Starter supports the following actions:

- **debugOn**: Turns on the debug mode for the bot. When the debug mode is on, the bot outputs to the user the details about the action plan to be executed and the time it took to execute this plan.
- **debugOff**: Turns off the debug mode for the bot. This will stop showing the action plan to be executed and will also stop showing the performance metrics for that action.
- **getSemanticInfo**: Retrieves the public information on the company found in the user's entry prompt. This action takes one mandatory parameter `entity`, which would be filled in by the result of running the `findEntity` Semantic Skill.
- **getCompanyDetails**: Retrieves the financial details about the `entity`, and as an example, augments that public data with a "proprietary" sample ESG scores from the custom action skill implementation. This can be replaced with your own business logic.
- **chatWithDocument**: This action summarizes or extracts key points from uploaded Text or PDF documents. It also allows the user to ask questions related to the uploaded documents.
- **webRetrieval**: When the user provides a web link with their prompt, this action extracts relevant information from the webpage's textual content. It also allows a user to follow up on additional questions related to this web content until the user explicitly asks to forget that document.
- **forgetDocuments**: Forgets all uploaded documents or website content for the current conversation.
