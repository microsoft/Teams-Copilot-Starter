import "reflect-metadata";

import { Utils } from "../../../src/helpers/utils"; // adjust the import path to your actual file structure
import fetch from "node-fetch";
jest.mock("node-fetch");
jest.mock("../../../src/env");

const publicFileName =
  "https://tcspublicfiles.blob.core.windows.net/files/w.txt";
const publicFileNameNotExist =
  "https://tcspublicfiles.blob.core.windows.net/files/notexist.txt";

describe("Utils", () => {
  describe("downloadFile", () => {
    it("should download file and return content", async () => {
      const mockFetch = fetch as jest.MockedFunction<typeof fetch>;
      const mockResponse = {
        ok: true,
        text: jest.fn().mockResolvedValue("file content"),
      };
      mockFetch.mockResolvedValue(mockResponse as any);

      const content = await Utils.downloadFile(publicFileName);

      expect(content).toBe("file content");
      expect(mockFetch).toHaveBeenCalledWith(publicFileName);
    });

    it("should throw an error if the response is not ok", async () => {
      const mockFetch = fetch as jest.MockedFunction<typeof fetch>;
      const mockResponse = {
        ok: false,
        status: 404,
        text: jest.fn().mockResolvedValue("Not Found"),
      };
      mockFetch.mockResolvedValue(mockResponse as any);

      await expect(Utils.downloadFile(publicFileNameNotExist)).rejects.toThrow(
        `Failed to download file from ${publicFileNameNotExist}. Status: 404`
      );
    });

    it("should throw an error if the fetch fails", async () => {
      const mockFetch = fetch as jest.MockedFunction<typeof fetch>;
      mockFetch.mockRejectedValue(new Error("Network error"));

      await expect(Utils.downloadFile(publicFileNameNotExist)).rejects.toThrow(
        "Network error"
      );
    });
  });
});
