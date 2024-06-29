import { ActivityTypes, Channels, TurnContext } from "botbuilder";
import { AI, Message, PredictedSayCommand } from "@microsoft/teams-ai";
import { Utils } from "../helpers/utils";
import { AIEntity } from "@microsoft/teams-ai/lib/actions";
import { ApplicationTurnState } from "../models/aiTypes";

/**
 * Formats the response from the AI and sends it to the user.
 * @param context The turn context.
 * @param state The application turn state.
 * @param command The predicted say command or message.
 * @param action The action to return.
 * @returns A promise that resolves to a string representing the action.
 */
export async function formatActionMessage(
  context: TurnContext,
  state: ApplicationTurnState,
  command: PredictedSayCommand | Message<string>,
  action?: string
): Promise<string> {
  const response = (command as PredictedSayCommand).response ?? command;
  if (!response?.content) {
    return action ?? AI.StopCommandName;
  }

  let content = response.content;
  const isTeamsChannel = context.activity.channelId === Channels.Msteams;

  if (isTeamsChannel) {
    content = content.split("\n").join("<br>");
  }

  // If the response from AI includes citations, they will be parsed and added to the response
  // eslint-disable-next-line prefer-const
  let [contentText, referencedCitations] =
    response.context && response.context.citations.length > 0
      ? Utils.formatCitations(content, response.context.citations)
      : [response.content, null];

  if (referencedCitations && referencedCitations.length > 0) {
    contentText += `<br><br> ⬇️ ${referencedCitations.length} references<br>`;

    referencedCitations.forEach((citation) => {
      contentText += `${citation.position}: ${citation.appearance.name}<br>`;
    });
  }

  // Send the response
  await context.sendActivity({
    type: ActivityTypes.Message,
    text: contentText,
    ...(isTeamsChannel ? { channelData: { feedbackLoopEnabled: true } } : {}),
    entities: [
      {
        type: "https://schema.org/Message",
        "@type": "Message",
        "@context": "https://schema.org",
        "@id": "",
        additionalType: ["AIGeneratedContent"],
        ...(referencedCitations ? { citation: referencedCitations } : {}),
      },
    ] as AIEntity[],
  });

  return action ?? AI.StopCommandName;
}
