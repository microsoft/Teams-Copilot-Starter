import "reflect-metadata";
import {
  ActionPlanner,
  Citation,
  ClientCitation,
  PromptTemplate,
} from "@microsoft/teams-ai";
import { AllowedFileTypes, ApplicationTurnState } from "../models/aiTypes";
import { Attachment, TurnContext } from "botbuilder";
import path from "path";
import fs from "fs";
import * as mime from "mime-types";
import * as responses from "../resources/responses";
import {
  EntityInfoSkill,
  GeneratePromptsSkill,
  SimilarEntitiesSkill,
} from "../skills";
import { Utils } from "./utils";
import CompanyInfo from "../models/companyInfo";
import EntityInfo from "../models/entityInfo";
import companyInfoCard from "../adaptiveCards/templates/companyInfo.json";
import { FileAttachment } from "../models/fileAttachment";
import { logging } from "../telemetry/loggerManager";
import { Env } from "../env";
import { AzureOpenAIEmbeddingsOptions, OpenAIEmbeddingsOptions } from "vectra";
import { container } from "tsyringe";

const logger = logging.getLogger("bot.TeamsAI");

/**
 * A class that provides methods to set and get cache records.
 */
export class ActionsHelper {
  /**
   * Check if the user has uploaded a document and if so, add properties to state
   * @param context
   * @param state
   * @returns arrary of links to the uploaded documents
   */
  public static async checkForUploadedFile(
    context: TurnContext,
    state: ApplicationTurnState
  ): Promise<FileAttachment[] | undefined> {
    const docs: FileAttachment[] = [];

    const fileAttachments = context.activity.attachments?.filter(
      (attachment) =>
        attachment.contentType ===
        "application/vnd.microsoft.teams.file.download.info"
    );
    if ((fileAttachments?.length ?? 0) > 0) {
      // only handle allowed filetypes
      if (!this.checkFileTypes(fileAttachments)) {
        await context.sendActivity(`You uploaded a document with
        name : '${fileAttachments?.[0]?.name}'. I can only process text or pdf file types for now. Please upload the correct file types documents.`);
      }

      (fileAttachments ?? []).forEach((attachment) => {
        docs.push({
          fileName: attachment.name!,
          url: attachment.content.downloadUrl!,
          type: attachment.content.fileType.toLowerCase(),
        } as FileAttachment);
      });
      if (docs.length > 0) {
        // Send a message to the user
        await context.sendActivity(
          docs.length > 1
            ? `You have uploaded files: '${docs
                .map((doc) => doc.fileName)
                .join(", ")}'`
            : `You have uploaded file: '${docs[0].fileName}'`
        );
      }
    }

    // Save the uploaded documents to the state
    if (fileAttachments && fileAttachments.length > 0) {
      const attachments = fileAttachments.map((attachment) => ({
        // check for the PDF file type, if not, default to text/html
        type: (
          attachment.content
            ? attachment.content.fileType === "pdf"
            : mime.contentType(attachment.name!) === "application/pdf"
        )
          ? "pdf"
          : "text/html",
        url: attachment.content.downloadUrl!,
        fileName: attachment.name!,
      }));
      // add the uploaded documents to the state
      state.conversation.uploadedDocuments =
        state.conversation.uploadedDocuments ?? [];
      state.conversation.uploadedDocuments?.push(...attachments);
    }

    // Log the uploaded files
    state.conversation.uploadedDocuments?.forEach((doc) => {
      logger.info(`Uploaded document: ${doc.fileName} of type ${doc.type}`);
    });

    // Show the user a message that the documents are being processed
    await context.sendActivity(responses.processingUploadedDocuments());
    return state.conversation.uploadedDocuments;
  }

  /**
   * Check if the user has uploaded a document and if so, add properties to state
   * @param context
   * @param state
   * @returns arrary of links to the uploaded documents
   */
  public static checkFileTypes(fileAttachments: Attachment[] | undefined) {
    (fileAttachments ?? []).forEach((attachment) => {
      if (
        !AllowedFileTypes.includes(attachment.content.fileType.toLowerCase())
      ) {
        return false;
      }
    });
    return true;
  }

  /**
   * Generates Adaptive Card for the provided company name
   * @param context
   * @param state
   * @param data
   * @returns
   */
  public static async generateAdaptiveCardForEntity(
    context: TurnContext,
    state: ApplicationTurnState,
    entity: EntityInfo,
    planner: ActionPlanner<ApplicationTurnState>
  ): Promise<Attachment | undefined> {
    try {
      const entityName = entity.companyInfo.name;

      // Generate prompts to follow up on the provided company name using OpenAI's GPT-3.5 API.
      const generatePromptsSkill = new GeneratePromptsSkill(
        context,
        state,
        planner
      );

      // Call Entity Info Skill to enrich the entity details with ESG scores
      const entityDetailsSkill = new EntityInfoSkill(context, state, planner);

      // Run parallel tasks
      const [entityInfo, prompts] = await Promise.all([
        entityDetailsSkill.run(entity),
        generatePromptsSkill.run(entityName),
      ]);

      if (prompts && prompts.length > 0) {
        // Prompts found
        logger.info(`Prompts found for '${entityName}'`);
        entityInfo.prompts = prompts;
      } else {
        // No prompts found
        logger.info(`Prompts not found for '${entityName}'`);
      }

      // Find similar companies
      const findSimiliarSkill = new SimilarEntitiesSkill(
        context,
        state,
        planner
      );
      const otherCompanies = (await findSimiliarSkill.run(
        entityName
      )) as CompanyInfo[];

      if (otherCompanies) {
        otherCompanies.forEach((entity) => {
          logger.info(`Entity found: '${entity.name}'`);
        });
        entityInfo.otherCompanies = otherCompanies;
      }

      // Send an adaptive card with the details
      const card = Utils.renderAdaptiveCard(companyInfoCard, {
        entity: entityInfo,
      });
      return card;
    } catch (error: any) {
      logger.error(`Error generating adaptive card for entity: ${error}`);
      throw error;
    }
  }

  public static getEmbeddingsOptions() {
    const env = container.resolve<Env>(Env);

    switch (env.data.OPENAI_TYPE) {
      case "AzureOpenAI":
        return {
          azureApiKey: env.data.OPENAI_KEY,
          azureEndpoint: env.data.OPENAI_ENDPOINT,
          azureDeployment: env.data.OPENAI_EMBEDDING_MODEL,
          azureApiVersion: env.data.OPENAI_API_VERSION,
        } as AzureOpenAIEmbeddingsOptions;
        break;
      case "OpenAI":
        return {
          model: env.data.OPENAI_EMBEDDING_MODEL,
          apiKey: env.data.OPENAI_KEY,
        } as OpenAIEmbeddingsOptions;
        break;
      default:
        // If using a CustomAI, add implementation here
        throw new Error("CustomAI is not supported for embeddings");
        break;
    }
  }

  /**
   * Formats the citations from the AI response into a format that can be used by the client
   * @param content The content from the AI response
   * @param citations The citations from the AI response
   * @returns The formatted citations
   */
  public static formatCitations(citations: Citation[]): ClientCitation[] {
    // If the response from AI includes citations, they will be parsed and added to the response
    const clientCitations = citations.map((citation, i) => {
      return {
        "@type": "Claim",
        position: `${i + 1}`,
        appearance: {
          "@type": "DigitalDocument",
          name: citation.title,
          abstract: Utils.extractSnippet(citation.content, 500),
        },
      } as ClientCitation;
    });

    return clientCitations;
  }

  /**
   * Adds the Azure AI Search data source to the provided prompt template
   * @param promptTemplate The name of the prompt template
   * @param planner The action planner
   * @returns The updated prompt template
   */
  public static async addAzureAISearchDataSource(
    promptTemplate: string,
    planner: ActionPlanner<ApplicationTurnState>
  ): Promise<any> {
    // Get the prompts from the planner
    const prompts = planner.prompts;

    // Get the environment settings
    const env = container.resolve<Env>(Env);

    // Get the prompt template for the provided prompt folder
    const template = await prompts.getPrompt(promptTemplate);

    // Read the SKPrompt from the file
    // eslint-disable-next-line security/detect-non-literal-fs-filename
    const skprompt = fs.readFileSync(
      path.join(__dirname, "..", "prompts", promptTemplate, "skprompt.txt")
    );

    //
    // Use the Azure AI Search data source for RAG over documents
    //
    const dataSources = (template.config.completion as any)["data_sources"];

    if (dataSources && dataSources.length > 0 && env.data) {
      dataSources.forEach((dataSource: any) => {
        if (dataSource.type === "azure_search" && dataSource.parameters) {
          dataSource.parameters.endpoint = env.data.AZURE_SEARCH_ENDPOINT;
          dataSource.parameters.authentication.key = env.data.AZURE_SEARCH_KEY;
          dataSource.parameters.index_name = env.data.AZURE_SEARCH_INDEX_NAME;
          dataSource.parameters.role_information = `${skprompt.toString(
            "utf-8"
          )}`;
          if (dataSource.parameters.embedding_dependency) {
            dataSource.parameters.embedding_dependency.deployment_name =
              env.data.OPENAI_EMBEDDING_MODEL;
          }
        }
      });
    } else {
      logger.error("No data sources found in the environment settings.");
    }

    return dataSources;
  }
}
