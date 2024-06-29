import { TurnContext } from "botbuilder";
import { AI, ActionPlanner } from "@microsoft/teams-ai";
import { ApplicationTurnState } from "../models/aiTypes";
import { ChatGPTSkill } from "../skills";
import { Utils } from "../helpers/utils";
import { UserHelper } from "../helpers/userHelper";
import { logging } from "../telemetry/loggerManager";
import * as responses from "../resources/responses";
import { formatActionMessage } from "./formatter";

/**
 * Retrieves semantic generic information using the Chat GPT Skill.
 * @param context The turn context.
 * @param state The application turn state.
 * @param planner The action planner.
 * @returns A promise that resolves to a string representing the response from the Chat GPT Skill.
 */
export async function getSemanticInfo(
  context: TurnContext,
  state: ApplicationTurnState,
  planner: ActionPlanner<ApplicationTurnState>
): Promise<string> {
  const logger = logging.getLogger("bot.TeamsAI");
  // Show typing indicator
  await Utils.startTypingTimer(context, state);

  // Get the user's information
  const user = await UserHelper.updateUserInfo(context, state);

  // Get the user's message
  const input = context.activity.text;

  // call Chat GPT Skill to get the generic response
  const chatGPTSkill = new ChatGPTSkill(context, state, planner);

  // Run the skill
  const response = await chatGPTSkill.run(input);
  if (response) {
    logger.info(`Chat response sent: '${response.content}'`);

    // Send the formatted response that may include the reference document citations
    return await formatActionMessage(context, state, response);
  } else {
    // No adaptive card found
    logger.info(`No response from GPT has been generated for '${input}'`);
    await context.sendActivity(responses.promptNotFound());
    return AI.StopCommandName;
  }
}
