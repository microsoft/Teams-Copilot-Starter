// This test file testing the rendering of Adaptive Cards.
import "reflect-metadata";

import { Utils } from "../../../src/helpers/utils";
import welcome from "../../../src/adaptiveCards/templates/welcome.json";
jest.mock("../../../src/env");

describe("Render Adaptive Card", () => {
  it("should return a welcome adaptive card", () => {
    const card = Utils.renderAdaptiveCard(welcome);
    expect(JSON.stringify(card.content)).toBe(JSON.stringify(welcome));
  });
});
