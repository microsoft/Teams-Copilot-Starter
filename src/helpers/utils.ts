// Import necessary modules
import {
  CardFactory,
  Attachment,
  MessagingExtensionAttachment,
  CardAction,
  ActivityTypes,
  TurnContext,
} from "botbuilder";
import EntityInfo from "../models/entityInfo";
import CompanyInfo from "../models/companyInfo";
import fetch from "node-fetch";
import * as ACData from "adaptivecards-templating";
import {
  Citation,
  ClientCitation,
  PredictedCommand,
} from "@microsoft/teams-ai";
import { ApplicationTurnState } from "../models/aiTypes";
import { container } from "tsyringe";
import { Env } from "../env";
import { TeamsAI } from "../bot/teamsAI";

const TYPING_TIMER_DELAY = 1000;
// Define a Utils class
export class Utils {
  static async MaxTurnsToRemember(): Promise<number> {
    const env = container.resolve(Env);
    return env.data.MAX_TURNS! * 2;
  }

  /**
   * Ensures that the response is in JSON format.
   * @param content The content to check and possibly format.
   * @returns The content in JSON format if wrapped in markdown, otherwise returns the original content.
   */
  static ensureJsonResponse(content: string): any {
    const regex = /^```json([\s\S]*?)```$/;
    const match = regex.exec(content);
    if (match && match.length > 1) {
      const jsonContent = match[1].trim();
      return JSON.parse(jsonContent);
    }
    try {
      // Try to parse the original content as JSON
      return JSON.parse(content);
    } catch (error) {
      return content;
    }
  }

  // Function to remove the JSON markdown from the response content
  static extractJsonResponse(content?: string): string {
    if (!content) {
      return "";
    }

    try {
      // Parse the content to check if it is an Action Plan JSON string
      const actionPlanJson = Utils.ensureJsonResponse(content);
      if (actionPlanJson && actionPlanJson?.commands) {
        return actionPlanJson?.commands.find((cmd: any) => cmd.type === "SAY")
          ?.response;
      } else {
        return actionPlanJson?.content ?? content;
      }
    } catch (error) {
      return content;
    }
  }

  /**
   * Binds an AdaptiveCard template with data.
   * @param rawCardTemplate The raw adaptive card template.
   * @param data The data to bind with the adaptive card template.
   * @returns An attachment containing the adaptive card with the bound data.
   */
  static renderAdaptiveCard(rawCardTemplate: any, data?: any): Attachment {
    // Create a new adaptive card template from the raw card template
    const cardTemplate = new ACData.Template(rawCardTemplate);
    // Expand the card template with the provided data
    const cardWithData = cardTemplate.expand({
      $root: data,
    });
    // Create a new adaptive card with the expanded data
    const card = CardFactory.adaptiveCard(cardWithData);
    // Return the adaptive card as an attachment
    return card;
  }

  static renderMessageExtensionAttachment(
    rawCardTemplate: any,
    data?: any
  ): MessagingExtensionAttachment {
    // Create a new adaptive card template from the raw card template
    const cardTemplate = new ACData.Template(rawCardTemplate);
    // Expand the card template with the provided data
    const cardWithData = cardTemplate.expand({
      $root: data,
    });
    // Create a new adaptive card with the expanded data
    const card = CardFactory.adaptiveCard(cardWithData);
    // Return the adaptive card as an attachment
    return card;
  }

  /**
   * Create a Messaging Extension Search Result Card
   * @param company The company information
   * @returns A Messaging Extension Search Result Card
   */
  static createMessageExtensionSearchResultCard(
    company: CompanyInfo
  ): MessagingExtensionAttachment {
    // TODO: Use "Adaptive Card Hero Card" since Hero card is considered deprecated
    const card = CardFactory.heroCard(company.name, [], [], {
      text: company.address?.city_state,
      subtitle: company.ticker,
    }) as MessagingExtensionAttachment;
    // Set the tap action
    card.preview = CardFactory.heroCard(company.name, [], [], {
      tap: { type: "invoke", value: { entity: company } } as CardAction,
      subtitle: company.ticker,
      text: company.worldRegion,
    });
    return card;
  }

  static createSearchResultCard(result: any): MessagingExtensionAttachment {
    // TODO: Use "Adaptive Card Hero Card" since Hero card is considered deprecated
    const card = CardFactory.heroCard(result.name, [], [], {
      text: result,
    }) as MessagingExtensionAttachment;
    card.preview = CardFactory.heroCard(result.name, [], [], {
      tap: { type: "invoke", value: result } as CardAction,
      text: result,
    });
    return card;
  }

  // get an adaptive card with data bound to it
  static getAdaptiveCardWithData(
    template: ACData.Template,
    data: any
  ): Attachment {
    // Bind the data to the template
    const cardWithData = template.expand({
      $root: {
        ...data,
      },
    });

    // Convert the card to an adaptive card attachment
    return CardFactory.adaptiveCard(cardWithData);
  }

  static createM365SearchResultAdaptiveCard(
    company: CompanyInfo,
    botId: string
  ): Attachment {
    const cardTemplate = Utils.getCompaniesListAdaptiveCardTemplate();
    return Utils.getAdaptiveCardWithData(cardTemplate, {
      entity: company,
      handOffToBotUrl: `${TeamsAI.HandoffUrl.replace(
        "${continuation}",
        company.id
      )}`,
    });
  }

  static createM365SearchResultHeroCard(company: CompanyInfo): Attachment {
    const card = CardFactory.heroCard(company?.name, [], [], {
      text: "",
    });
    card.content.tap = {
      type: "invoke",
      value: {
        verb: "getSemanticInfo",
        entity: company,
      },
    };
    return card;
  }

  /**
   * Converts a string to PascalCase.
   * @param text The string to convert to PascalCase.
   * @returns The converted string in PascalCase.
   */
  static toPascalCase(text: string): string {
    return text
      .toLowerCase()
      .replace(/[-_]+/g, " ")
      .replace(/[^\w\s]/g, "")
      .replace(
        /\s+(.)(\w*)/g,
        (_, firstChar, rest) => `${firstChar.toUpperCase()}${rest}`
      )
      .replace(/\w/, (s) => s.toUpperCase());
  }

  /**
   * Checks if a record with the specified key exists in an array of records.
   * @param rec The record to check.
   * @param key The key to check for in the record.
   * @returns True if a record with the specified key exists, false otherwise.
   */
  static isRecordExists<T>(rec: Record<string, T>, key: string): boolean {
    // Check if rec is an object and has a property called "key"
    return key in rec;
  }

  /**
   * Checks if an entity has expired based on the specified expiration time.
   * @param entity The entity to check for expiration.
   * @param expirationTime The expiration time in seconds.
   * @returns True if the entity has expired, false otherwise.
   */
  static isEntityExpired(entity: EntityInfo, expirationTime: number): boolean {
    const now = new Date();
    const lastUpdated: Date = entity.lastUpdated
      ? new Date(entity.lastUpdated)
      : now;
    const diff = (now.getTime() - lastUpdated.getTime()) / 1000;
    return diff > expirationTime;
  }

  /**
   * Downloads the contents of a file from a URL.
   * @param url The URL to download the file from.
   * @returns The text content of the downloaded file.
   */
  static async downloadFile(url: string): Promise<string> {
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(
        `Failed to download file from ${url}. Status: ${response.status}`
      );
    }
    return await response.text();
  }

  /**
   * Creates a list of attachments with detailed information for each company.
   * @param companyNames List of company information.
   * @returns List of messaging extension attachments.
   */
  static async createCompanyListAttachments(
    companyNames: CompanyInfo[]
  ): Promise<MessagingExtensionAttachment[]> {
    const companyListAttachments: MessagingExtensionAttachment[] = [];

    companyNames?.forEach((company: CompanyInfo) =>
      companyListAttachments.push(
        Utils.createMessageExtensionSearchResultCard(company)
      )
    );

    return companyListAttachments;
  }

  static findFirstCommonWords(input1: string, input2: string): string {
    // Split the inputs into arrays of words
    const words1 = input1.toLowerCase().match(/\b\w+\b/g) || [];
    const words2 = input2.toLowerCase().match(/\b\w+\b/g) || [];

    // Find the common words
    const commonWords = words1.filter((word) => words2.includes(word as never));

    return commonWords?.length > 0
      ? commonWords[0].charAt(0).toUpperCase() + commonWords[0].slice(1)
      : "";
  }

  static getCompaniesListAdaptiveCardTemplate(): ACData.Template {
    // Deliverable adaptive card template
    const cardTemplate = new ACData.Template({
      type: "AdaptiveCard",
      body: [
        {
          type: "Container",
          items: [
            {
              type: "ColumnSet",
              columns: [
                {
                  type: "Column",
                  width: "auto",
                  items: [
                    {
                      type: "Image",
                      url: "${entity.logoUrl}",
                      size: "small",
                    },
                  ],
                },
                {
                  type: "Column",
                  width: "stretch",
                  items: [
                    {
                      type: "TextBlock",
                      text: "${entity.name}",
                      weight: "Bolder",
                      size: "Large",
                      spacing: "None",
                      wrap: true,
                      style: "heading",
                    },
                  ],
                },
              ],
            },
          ],
        },
        {
          type: "TextBlock",
          text: "Stock Market: ${if(empty([entity.ticker]), 'N/A', [entity.ticker])}",
          spacing: "None",
          size: "Small",
          wrap: true,
          isSubtle: true,
        },
        {
          type: "TextBlock",
          text: "Location: ${entity.worldRegion}",
          wrap: true,
          style: "heading",
        },
        {
          type: "TextBlock",
          text: "Website: ${entity.website}",
          wrap: true,
          style: "heading",
        },
      ],
      actions: [
        {
          type: "Action.Submit",
          title: "View Details",
          data: {
            verb: "getSemanticInfo",
            msteams: {
              type: "task/fetch",
            },
            command: "getSemanticInfo",
            entity: "${entity}",
          },
        },
        {
          type: "Action.OpenUrl",
          title: "Handoff to Teams Copilot",
          url: "${handOffToBotUrl}",
        },
      ],
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      version: "1.4",
    });

    return cardTemplate;
  }

  // Function to extract the text response from the Plan Action Command of type "SAY"
  static convertActionPlanSayResponseToText(actionPlan: string): string {
    try {
      const actionPlanJson = JSON.parse(actionPlan);
      return actionPlanJson?.commands.find((cmd: any) => cmd.type === "SAY")
        ?.response;
    } catch (error) {
      return actionPlan;
    }
  }

  // Function to swap "DO" and "SAY" commands
  static swapDoAndSay(commands: PredictedCommand[]): PredictedCommand[] {
    for (let i = 0; i < commands.length; i++) {
      if (commands[i].type === "DO" && commands[i + 1]?.type === "SAY") {
        const temp = commands[i];
        commands[i] = commands[i + 1];
        commands[i + 1] = temp;
      }
    }
    return commands;
  }

  /**
   * Ensures that `key` is a properly sanitized Azure Blob Storage key. It should be URI encoded,
   * no longer than 1024 characters, and contain no more than 254 slash ("/") chars.
   *
   * @param {string} key string blob key to sanitize
   * @returns {string} sanitized blob key
   */
  static sanitizeBlobKey(key: string): string {
    if (!key || !key.length) {
      throw new Error("Please provide a non-empty key");
    }

    const sanitized = key
      .split("/")
      .reduce((acc, part, idx) =>
        part ? [acc, part].join(idx < 255 ? "/" : "") : acc
      );
    return encodeURIComponent(sanitized).substring(0, 1024);
  }

  /**
   * Extract snippets from the response content limited to the given length.
   * @param {string} content The content to extract snippets from.
   * @param {number} length The maximum length of the snippets.
   * @returns {string} The extracted snippets.
   */
  static extractSnippet(content: string, length: number): string {
    const snippets = content.split("\n").filter((line) => line.length > 0);
    let snippet = "";
    let remaining = length;
    for (const line of snippets) {
      if (line.length < remaining) {
        snippet += line + "\n";
        remaining -= line.length + 1;
      } else {
        snippet += line.substring(0, remaining) + "\n";
        break;
      }
    }
    return snippet;
  }

  /**
   * Format the response content so that the sources are numbered instead of [doc1].
   * @param {string} content The content to format.
   * @returns {string} The formatted content.
   */
  static formatCitationsResponse(content: string): string {
    // Replace all occurrences of `[docX]` with `[X]`
    const newContent = content.replace(/\[doc(\d+)\]/g, (match, p1) => {
      return `[${p1}]`;
    });

    return newContent;
  }

  /**
   * Manually start a timer to periodically send "typing" activities.
   * @remarks
   * The timer waits 1000ms to send its initial "typing" activity and then send an additional
   * "typing" activity every 1000ms. The timer will automatically end once an outgoing activity
   * has been sent. If the timer is already running or the current activity, is not a "message"
   * the call is ignored.
   * @param {TurnContext} context The context for the current turn with the user.
   */
  public static async startTypingTimer(
    context: TurnContext,
    state: ApplicationTurnState
  ): Promise<void> {
    if (
      context.activity.type == ActivityTypes.Message &&
      !state.temp.typingTimer
    ) {
      // Listen for outgoing activities
      context.onSendActivities((context, activities, next) => {
        // Listen for any messages to be sent from the bot
        if (timerRunning) {
          for (let i = 0; i < activities.length; i++) {
            // TODO:
            // eslint-disable-next-line security/detect-object-injection
            if (activities[i].type == ActivityTypes.Message) {
              // Stop the timer
              this.stopTypingTimer(state);
              timerRunning = false;
              break;
            }
          }
        }
        return next();
      });
      let timerRunning = true;
      const onTimeout = async () => {
        try {
          // Send typing activity
          await context.sendActivity({ type: ActivityTypes.Typing });
        } catch (err) {
          // Seeing a random proxy violation error from the context object. This is because
          // we're in the middle of sending an activity on a background thread when the turn ends.
          // The context object throws when we try to update "this.responded = true". We can just
          // eat the error but lets make sure our states cleaned up a bit.
          state.temp.typingTimer = undefined;
          timerRunning = false;
        }
        // Restart timer
        if (timerRunning) {
          state.temp.typingTimer = setTimeout(onTimeout, TYPING_TIMER_DELAY);
        }
      };
      state.temp.typingTimer = setTimeout(onTimeout, TYPING_TIMER_DELAY);
    }
  }

  /**
   * Manually stop the typing timer.
   * @remarks
   * If the timer isn't running nothing happens.
   */
  public static async stopTypingTimer(
    state?: ApplicationTurnState
  ): Promise<void> {
    if (state?.temp) {
      clearTimeout(state.temp.typingTimer);
      state.temp.typingTimer = undefined;
    }
  }

  /**
   * Extracts citations from the content using a regular expression that uses [doc#] to create citations references.
   * @param {string} content The content to extract citations from.
   * @param {Citation[]} contextCitations The citations to extract from the content.
   * @returns {string, Citation[]} The formatted content and extracted citations in form of a tuple.
   */
  public static formatCitations(
    content: string,
    contextCitations: Citation[]
  ): [string, ClientCitation[] | undefined] {
    const env = container.resolve(Env);

    // If the response from AI includes citations, they will be parsed and added to the response
    const citations = contextCitations.map((citation, i) => {
      return {
        "@type": "Claim",
        position: `${i + 1}`,
        appearance: {
          "@type": "DigitalDocument",
          name: citation.title,
          abstract: Utils.extractSnippet(citation.content, 500),
          url: `${TeamsAI.HandoffUrl.replace(
            "${continuation}",
            citation.url ?? citation.title ?? ""
          )}`,
          usageInfo: {
            type: "https://schema.org/Message",
            "@type": "CreativeWork",
            name: "Confidentiality Policy",
            description:
              "This document is confidential and should not be shared with unauthorized personnel.",
          },
        },
      } as ClientCitation;
    });

    // If there are citations, modify the content so that the sources are numbered instead of [doc1]
    const contentText = !citations
      ? content
      : Utils.formatCitationsResponse(content);

    // If there are citations, filter out the citations unused in content.
    const referencedCitations = citations
      ? Utils.filterUnusedCitations(contentText, citations)
      : undefined;

    return [contentText, referencedCitations];
  }

  /**
   * Filters out citations that are not used in the content.
   * @param {string} contentText The content text to filter citations from.
   * @param {ClientCitation[]} citations The citations to filter.
   * @returns {ClientCitation[]} The filtered citations.
   */
  public static filterUnusedCitations(
    contentText: string,
    citations: ClientCitation[]
  ): ClientCitation[] {
    return citations.filter((citation) =>
      contentText.includes(`[${citation.position}]`)
    );
  }
}
