import { Channels, TurnContext } from "botbuilder";
import { ActionPlanner } from "@microsoft/teams-ai";
import { ApplicationTurnState } from "../../models/aiTypes";
import { ChatGPTSkill } from "../../skills";
import { Utils } from "../../helpers/utils";
import { logging } from "../../telemetry/loggerManager";
import * as responses from "../../resources/responses";
import copilotCard from "../templates/copilotResponse.json";

/**
 * Sends a suggested prompt to the user and displays the response using an Adaptive Card.
 * @param {TurnContext} context - The context object for the current turn of the conversation.
 * @param {ApplicationTurnState} state - The application turn state.
 * @param {any} data - The prompt message data.
 * @param {ActionPlanner<ApplicationTurnState>} planner - The action planner.
 * @returns {Promise<string>} A promise that resolves to a string indicating the result of the operation.
 */
export async function suggestedPrompt(
  context: TurnContext,
  state: ApplicationTurnState,
  data: any,
  planner: ActionPlanner<ApplicationTurnState>
): Promise<string> {
  const logger = logging.getLogger("bot.TeamsAI");

  // notify user that the bot is working
  await context.sendActivity(responses.promptReturned());

  // Show typing indicator
  await Utils.startTypingTimer(context, state);

  // call Chat GPT Skill to get the generic response
  const chatGPTSkill = new ChatGPTSkill(context, state, planner);

  // Run the skill
  const promptResponse = await chatGPTSkill.run(data as string);
  if (!promptResponse) {
    // No prompt response found
    logger.info(`Prompt response not found for '${data}'`);
    await context.sendActivity(responses.promptNotFound());
    return "";
  }

  // If the response from AI includes citations, they will be parsed and added to the response
  const [contentText, referencedCitations] =
    promptResponse.context && promptResponse.context.citations.length > 0
      ? Utils.formatCitations(promptResponse, promptResponse.context.citations)
      : [promptResponse.content, null];

  // Send Adaptive Card with the prompt response
  const card = Utils.renderAdaptiveCard(copilotCard, {
    prompt: {
      request: data,
      response: contentText,
    },
  });

  // Render the Adaptive Card based on the retrieved company details
  // Render the Adaptive Card based on the retrieved company details
  await context.sendActivity({
    attachments: [card],
    ...(context.activity.channelId === Channels.Msteams
      ? { channelData: { feedbackLoopEnabled: true } }
      : {}),
    entities: [
      {
        type: "https://schema.org/Message",
        "@type": "Message",
        "@context": "https://schema.org",
        "@id": "",
        additionalType: ["AIGeneratedContent"],
        usageInfo: {
          name: "Confidential",
          description:
            "This message is confidential and intended only for internal use.",
        },
        ...(referencedCitations ? { citations: referencedCitations } : {}),
      },
    ],
  });

  // Return the result of the operation
  return "";
}
