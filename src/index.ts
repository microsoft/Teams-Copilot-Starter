// Import required packages
import "reflect-metadata";
import * as restify from "restify";
import * as path from "path";

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import {
  ConfigurationServiceClientCredentialFactory,
  TurnContext,
} from "botbuilder";

import { AxiosError } from "axios";
import { container } from "tsyringe";
import { logging } from "./telemetry/loggerManager";
import { configureTeamsAI } from "./configTeamsAI";
import { Env } from "./env";
import { ConsoleLogger } from "./telemetry/consoleLogger";
import { AppInsightLogger } from "./telemetry/appInsightLogger";
import { BlobsStorageLeaseManager } from "./helpers/blobsStorageLeaseManager";
import { TeamsAdapter } from "@microsoft/teams-ai";
import * as jwtValidator from "./services/jwtValidator";
import { getTickerQuote } from "./api/apiTicker";
import { BlobsStorage } from "botbuilder-azure-blobs";

// Create an instance of the environment variables
const envVariables: Env = new Env();

// Configure logging
const consoleLogger = new ConsoleLogger();
const appInsightLogger = new AppInsightLogger();

logging
  .configure({
    minLevels: {
      "": "trace",
    },
  })
  .registerLogger(consoleLogger)
  .registerLogger(appInsightLogger);

// Get logging
const logger = logging.getLogger("index");

// register the environment variables
container.register<Env>(Env, {
  useValue: envVariables,
});

// Create adapter.
const adapter = new TeamsAdapter(
  {},
  new ConfigurationServiceClientCredentialFactory({
    MicrosoftAppId: envVariables.data.BOT_ID,
    MicrosoftAppPassword: envVariables.data.BOT_PASSWORD,
    MicrosoftAppType: envVariables.data.BOT_APP_TYPE,
  })
);

// Due to bug in teams-ai, sso does not work correctly with BlobsStorage. T
// The method onUserSignInSuccess in the TeamsAdapter class is not called when using BlobsStorage.
// Therefore, MemoryStorage is used instead for the SSO use case.
// See: https://github.com/microsoft/teams-ai/issues/1457

// NOTE: Comment out the following lines (68-72) to use MemoryStorage instead of BlobsStorage to enable SSO for now
logger.info("Creating BlobsStorage");
const storage = new BlobsStorage(
  `DefaultEndpointsProtocol=https;AccountName=${envVariables.data.STORAGE_ACCOUNT_NAME};AccountKey=${envVariables.data.STORAGE_ACCOUNT_KEY};EndpointSuffix=core.windows.net`,
  envVariables.data.STORAGE_CONTAINER_NAME!
);

// NOTE: Uncomment the following line to use MemoryStorage instead of BlobsStorage to enable SSO for now
// const storage = new MemoryStorage();

const storageLeaseManager = new BlobsStorageLeaseManager(
  `DefaultEndpointsProtocol=https;AccountName=${envVariables.data.STORAGE_ACCOUNT_NAME};AccountKey=${envVariables.data.STORAGE_ACCOUNT_KEY};EndpointSuffix=core.windows.net`,
  `${envVariables.data.STORAGE_CONTAINER_NAME!}-state-manager`
);

container.register<BlobsStorageLeaseManager>(BlobsStorageLeaseManager, {
  useValue: storageLeaseManager,
});

// Create the bot that will handle incoming messages.
const bot = configureTeamsAI(storage, adapter, logger, envVariables);

// Add a custom response formatter to convert markdown code blocks to <pre> tags
// addResponseFormatter(bot.app);

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  logger.error(
    `[onTurnError] unhandled error: ${error}, stack: ${error.stack}`
  );

  if (error.name === "AxiosError") {
    const msg =
      `Network Error ${(<AxiosError<any>>error).response?.status}:` +
        (<AxiosError<any>>error).response?.data?.error?.message ??
      (<AxiosError<any>>error).response?.statusText ??
      "unknown error";
    logger.warn(`[onTurnError] error details: ${msg}`);
    await context.sendActivity(msg);
  } else {
    logger.error(`[onTurnError] unhandled error: ${error}`);
  }

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  );

  // Send a message to the user
  await context.sendActivity(
    "There was an error generating your response. Please try again. If the error persists, please contact your support team."
  );
};

// Set the onTurnError for the singleton CloudAdapter
adapter.onTurnError = onTurnErrorHandler;

logger.info("Starting bot");

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.listen(process.env.PORT || process.env.port || 3978, () => {
  logger.info(`Bot Started, ${server.name} listening to ${server.url}`);
});

// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
  await adapter
    .process(req, res, async (context) => {
      // If the bot is not running, start it.
      await bot.start(context);
      // Run the bot's message processing pipeline.
      await bot.app.run(context);
    })
    .catch((err) => {
      // Error message including "412" means it is waiting for user's consent, which is a normal process of SSO, sholdn't throw this error.
      if (!err.message.includes("412")) {
        throw err;
      }
    });
});

// Listen for incoming requests to get Ticker.
// This is a sample API that returns a random quote for a given ticker symbol.
// The API is protected by a JWT token. The token is validated by the jwtValidator middleware.
server.get("/api/quotes/:ticker", jwtValidator.validateJwt, getTickerQuote);

server.get(
  "/auth-:name(start|end).html",
  restify.plugins.serveStatic({
    directory: path.join(__dirname, "public"),
  })
);
