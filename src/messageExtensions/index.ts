import "reflect-metadata";
import { container } from "tsyringe";

import { CustomDataSourceSkill } from "../skills/customDataSourceSkill";
import companyListCard from "../adaptiveCards/templates/companyList.json";
import genHandoffCard from "../adaptiveCards/templates/genericHandoffCard.json";

import {
  Attachment,
  CardFactory,
  Channels,
  MessagingExtensionAttachment,
  TurnContext,
} from "botbuilder";
import { ActionPlanner, Query } from "@microsoft/teams-ai";
import { ApplicationTurnState, TData } from "../models/aiTypes";
import { Logger } from "../telemetry/logger";
import CompanyInfo from "../models/companyInfo";
import { TeamsAI } from "../bot/teamsAI";
import { Utils } from "../helpers/utils";
import { EntityRecognitionSkill } from "../skills";
import { Env } from "../env";
import * as querystring from "querystring";
import * as ACData from "adaptivecards-templating";
import axios from "axios";

/**
 * Searches for entities based on a query and returns the results as a list of attachments.
 * @param context The turn context.
 * @param state The application turn state.
 * @param query The query parameters.
 * @param planner The action planner.
 * @param logger The logger.
 * @returns A promise that resolves to the search results as a list of attachments.
 */
export async function searchCmd(
  context: TurnContext,
  state: ApplicationTurnState,
  query: Query<Record<string, any>>,
  planner: ActionPlanner<ApplicationTurnState>,
  logger: Logger
): Promise<any> {
  const searchQuery = query.parameters.companyName ?? "";

  if (
    context.activity.channelData.source.name === TeamsAI.M365CopilotSourceName
  ) {
    // call Entity Info Skill to get the entity details from Teams Copilot Starter API
    const entityRecognitionSkill = new EntityRecognitionSkill(
      context,
      state,
      planner
    );
    // Run the entity recognition skill to try and find the company name
    const listOfCompanies = await entityRecognitionSkill.findMatchingCompanies(
      searchQuery
    );

    if (!listOfCompanies || listOfCompanies.length === 0) {
      return {
        type: "message",
        text: "No entities found.",
      };
    }
    const env = container.resolve<Env>(Env);
    const botId = env.data.BOT_ID ?? "";

    // Render the Adaptive Card based on the retrieved company details
    const attachmentsPromise: Promise<MessagingExtensionAttachment>[] =
      listOfCompanies.map(async (company: CompanyInfo) => {
        logger.info(JSON.stringify(company));

        // create an adaptive card to be returned to M365 Copilot
        const card = Utils.createM365SearchResultAdaptiveCard(company, botId);

        // Create a preview card
        const preview = Utils.createM365SearchResultHeroCard(company);

        // Return the full attachment
        return { ...card, preview };
      });

    const attachments = await Promise.all(attachmentsPromise);
    logger.info(`Returning the list of ${listOfCompanies.length} items`);

    // Return results as a list
    return {
      type: "result",
      attachmentLayout: "list",
      attachments: attachments,
    };
  } else {
    // call Entity Recognition Skill to extract the entity name from the user's message
    const customDataSourceSkill = new CustomDataSourceSkill(
      context,
      state,
      planner
    );

    // Run the skill
    try {
      const foundEntities = (await customDataSourceSkill.run(
        searchQuery
      )) as CompanyInfo[];

      // if no entities were found, return a message to the user
      if (!foundEntities?.length ?? 0 === 0) {
        return {
          type: "message",
          text: "No entities found.",
        };
      }

      // call Entity Info Skill to get the entity details list
      const attachments = await Utils.createCompanyListAttachments(
        foundEntities
      );

      // log items returned
      logger.info(`Found ${attachments.length} items for '${searchQuery}'`);

      // Return results as a list
      return {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      };
    } catch (error) {
      return {
        type: "message",
        text: "An error occurred while processing the request.",
      };
    }
  }
}

/**
 * Selects an item and generates detailed information for the selected company.
 * @param context The turn context.
 * @param state The application turn state.
 * @param data The data containing the selected company entity.
 * @returns A promise that resolves to the detailed information as an attachment.
 */
export async function selectItem(
  context: TurnContext,
  state: ApplicationTurnState,
  data: TData
): Promise<any> {
  // Generate detailed information for the selected company
  const entity: CompanyInfo = data.entity;

  const card = Utils.renderAdaptiveCard(companyListCard, entity);

  if (context.activity.conversation.conversationType === "personal") {
    if (card) {
      await context.sendActivity({
        attachments: [card],
        ...(context.activity.channelId === Channels.Msteams
          ? { channelData: { feedbackLoopEnabled: true } }
          : {}),
        entities: [
          {
            type: "https://schema.org/Message",
            "@type": "Message",
            "@context": "http://schema.org",
            "@id": "",
            usageInfo: {
              name: "Confidential",
              description:
                "This message is confidential and intended only for internal use.",
            },
          },
        ],
      });
    } else {
      await context.sendActivity("No company information found.");
    }
  } else {
    return {
      attachmentLayout: "list",
      attachments: card ? [card] : [],
      type: "result",
    };
  }
}

/**
 * Searches for an npm package based on the query and returns the results as a list of attachments.
 * @param context
 * @param state
 * @param query
 * @param planner
 * @param logger
 */
export async function findNpmPackage(
  context: TurnContext,
  state: ApplicationTurnState,
  query: Query<Record<string, any>>,
  env: Env,
  logger: Logger
): Promise<any> {
  logger.info(`Query received: ${query.parameters.npmPackageName}`);
  const searchQuery = query.parameters.npmPackageName;
  const response = await axios.get(
    `http://registry.npmjs.com/-/v1/search?${querystring.stringify({
      text: searchQuery,
      size: 8,
    })}`
  );

  const attachments: Attachment[] = [];
  response.data.objects.forEach((obj: any) => {
    const template = new ACData.Template(genHandoffCard);
    const card = template.expand({
      $root: {
        name: obj.package.name,
        description: obj.package.description,
        handoffUrl: `https://teams.microsoft.com/l/chat/0/0?users=28:${env.data.BOT_ID}&continuation=${obj.package.name}`,
      },
    });
    const preview = CardFactory.heroCard(obj.package.name);
    const attachment = { ...CardFactory.adaptiveCard(card), preview };
    attachments.push(attachment);
  });

  return {
    type: "result",
    attachmentLayout: "list",
    attachments: attachments,
  };
}
