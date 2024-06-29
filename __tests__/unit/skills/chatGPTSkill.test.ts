import "reflect-metadata";
import { ApplicationTurnState } from "../../../src/models/aiTypes";
import { ChatGPTSkill } from "../../../src/skills/chatGPTSkill"; // adjust the import path to your actual file structure
import { TurnContext } from "botbuilder";
import { ActionPlanner, PromptResponse } from "@microsoft/teams-ai";
import { Logger } from "../../../src/telemetry/logger";
import EventEmitter from "events";
import { Utils } from "../../../src/helpers/utils";

jest.mock("../../../src/telemetry/logger");
jest.mock("events");
jest.mock("../../../src/env");
jest.mock("../../../src/helpers/utils");

describe("ChatGPTSkill", () => {
  let chatGPTSkill: ChatGPTSkill;
  let mockTurnContext: jest.Mocked<TurnContext>;
  let mockApplicationTurnState: jest.Mocked<ApplicationTurnState>;
  let mockActionPlanner: jest.Mocked<ActionPlanner<ApplicationTurnState>>;
  let mockLogger: jest.Mocked<Logger>;
  let mockLogManager: jest.Mocked<EventEmitter>;

  beforeEach(() => {
    // Create a mock TurnContext
    mockTurnContext = {
      sendActivity: jest.fn(),
      // Add other methods as needed
    } as unknown as jest.Mocked<TurnContext>;

    // Create a mock ApplicationTurnState
    mockApplicationTurnState = {
      temp: {
        input: "test input",
      },
    } as unknown as jest.Mocked<ApplicationTurnState>;

    // Create a mock ActionPlanner
    mockActionPlanner = {
      // Mock the properties and methods of ActionPlanner as needed
    } as unknown as jest.Mocked<ActionPlanner<ApplicationTurnState>>;

    mockLogManager = new EventEmitter() as jest.Mocked<EventEmitter>;

    mockLogger = new Logger(
      mockLogManager,
      "module",
      "minLevel"
    ) as jest.Mocked<Logger>;

    chatGPTSkill = new ChatGPTSkill(
      mockTurnContext,
      mockApplicationTurnState,
      mockActionPlanner
    );
  });

  it("should return the expected result when run on chatGPTSkill is succesfully called", async () => {
    const input = "test input";
    const expectedContent = "expected result";
    Utils.MaxTurnsToRemember = jest.fn().mockResolvedValue(1);

    const mockedResult: PromptResponse<unknown> = {
      status: "success",
      message: { role: "assistant", content: "expected result" },
      input: { role: "user", content: input },
    };
    const mockCompletePrompt = jest.spyOn(
      ActionPlanner.prototype,
      "completePrompt"
    );
    mockCompletePrompt.mockImplementation(() => Promise.resolve(mockedResult));

    const mockedExtractJsonResponse = jest.spyOn(Utils, "extractJsonResponse");
    mockedExtractJsonResponse.mockImplementation(
      (inputString: string | undefined) => inputString ?? ""
    );

    const result = await chatGPTSkill.run(input);

    expect(result.content).toBe(expectedContent);
  });

  it("should return undefined when run on chatGPTSkill is not succesfully called", async () => {
    const input = "test input";
    Utils.MaxTurnsToRemember = jest.fn().mockResolvedValue(1);

    const mockedResult: PromptResponse<unknown> = {
      status: "error",
      input: { role: "user", content: input },
    };
    const mockCompletePrompt = jest.spyOn(
      ActionPlanner.prototype,
      "completePrompt"
    );
    mockCompletePrompt.mockImplementation(() => Promise.resolve(mockedResult));

    const result = await chatGPTSkill.run(input);

    expect(result).toBeUndefined();
    expect(mockTurnContext.sendActivity).toHaveBeenCalled();
  });

  it("should return undefined when run on chatGPTSkill has general Error", async () => {
    const input = "test input";
    Utils.MaxTurnsToRemember = jest.fn().mockResolvedValue(1);

    const mockedResult: PromptResponse<unknown> = {
      status: "error",
      error: {
        name: "General Error",
        message: "General error message",
      },
      input: { role: "user", content: input },
    };
    const mockCompletePrompt = jest.spyOn(
      ActionPlanner.prototype,
      "completePrompt"
    );
    mockCompletePrompt.mockImplementation(() => Promise.resolve(mockedResult));

    const result = await chatGPTSkill.run(input);

    expect(result).toBeUndefined();
    expect(mockTurnContext.sendActivity).toHaveBeenCalled();
  });

  it("should return undefined when run on chatGPTSkill has rate limited error", async () => {
    const input = "test input";
    Utils.MaxTurnsToRemember = jest.fn().mockResolvedValue(1);

    const mockedResult: PromptResponse<unknown> = {
      status: "rate_limited",
      error: {
        name: "rate_limited_error",
        message: "rate limited error message",
      },
      input: { role: "user", content: input },
    };
    const mockCompletePrompt = jest.spyOn(
      ActionPlanner.prototype,
      "completePrompt"
    );
    mockCompletePrompt.mockImplementation(() => Promise.resolve(mockedResult));

    const result = await chatGPTSkill.run(input);

    expect(result).toBeUndefined();
    expect(mockTurnContext.sendActivity).toHaveBeenCalled();
  });

  it("should return undefined when run on chatGPTSkill has invalid response error", async () => {
    const input = "test input";
    Utils.MaxTurnsToRemember = jest.fn().mockResolvedValue(1);

    const mockedResult: PromptResponse<unknown> = {
      status: "invalid_response",
      error: {
        name: "invalid_response_error",
        message: "invalid response error message",
      },
      input: { role: "user", content: input },
    };
    const mockCompletePrompt = jest.spyOn(
      ActionPlanner.prototype,
      "completePrompt"
    );
    mockCompletePrompt.mockImplementation(() => Promise.resolve(mockedResult));

    const result = await chatGPTSkill.run(input);

    expect(result).toBeUndefined();
    expect(mockTurnContext.sendActivity).toHaveBeenCalled();
  });
});
