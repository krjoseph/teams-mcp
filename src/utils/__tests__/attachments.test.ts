import { beforeEach, describe, expect, it, vi } from "vitest";
import type { GraphService } from "../../services/graph.js";
import {
  getFileExtensionFromMimeType,
  imageUrlToBase64,
  isValidImageType,
  uploadImageAsHostedContent,
} from "../attachments.js";

const mockGraphService = {
  getClient: vi.fn(),
} as unknown as GraphService;

const mockClient = {
  api: vi.fn(),
};

// Mock fetch globally
const mockFetch = vi.fn();
global.fetch = mockFetch;

describe("Attachment Utilities", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    (mockGraphService.getClient as any).mockResolvedValue(mockClient);
  });

  describe("uploadImageAsHostedContent", () => {
    it("should upload image data successfully", async () => {
      const mockResponse = { id: "hosted-content-123" };

      mockClient.api.mockReturnValue({
        post: vi.fn().mockResolvedValue(mockResponse),
      });

      const imageData =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";
      const result = await uploadImageAsHostedContent(
        mockGraphService,
        "team123",
        "channel456",
        imageData,
        "image/png",
        "test.png"
      );

      expect(result).toEqual({
        hostedContentId: "hosted-content-123",
        attachment: {
          id: "hosted-content-123",
          contentType: "image/png",
          contentUrl:
            "https://graph.microsoft.com/v1.0/teams/team123/channels/channel456/messages/hostedContents/hosted-content-123/$value",
          name: "test.png",
        },
      });

      expect(mockClient.api).toHaveBeenCalledWith(
        "/teams/team123/channels/channel456/messages/hostedContents"
      );
    });

    it("should upload buffer data successfully", async () => {
      const mockResponse = { id: "hosted-content-123" };

      mockClient.api.mockReturnValue({
        post: vi.fn().mockResolvedValue(mockResponse),
      });

      const imageBuffer = Buffer.from("test image data");
      const result = await uploadImageAsHostedContent(
        mockGraphService,
        "team123",
        "channel456",
        imageBuffer,
        "image/jpeg"
      );

      expect(result).toEqual({
        hostedContentId: "hosted-content-123",
        attachment: {
          id: "hosted-content-123",
          contentType: "image/jpeg",
          contentUrl:
            "https://graph.microsoft.com/v1.0/teams/team123/channels/channel456/messages/hostedContents/hosted-content-123/$value",
          name: "image.jpg",
        },
      });
    });

    it("should handle upload errors", async () => {
      mockClient.api.mockReturnValue({
        post: vi.fn().mockRejectedValue(new Error("Upload failed")),
      });

      const consoleSpy = vi.spyOn(console, "error").mockImplementation(() => {
        // Mock implementation - do nothing
      });

      const result = await uploadImageAsHostedContent(
        mockGraphService,
        "team123",
        "channel456",
        "imagedata",
        "image/png"
      );

      expect(result).toBeNull();
      expect(consoleSpy).toHaveBeenCalledWith(
        "Error uploading image as hosted content:",
        expect.any(Error)
      );

      consoleSpy.mockRestore();
    });
  });

  describe("isValidImageType", () => {
    it("should validate common image types", () => {
      expect(isValidImageType("image/jpeg")).toBe(true);
      expect(isValidImageType("image/jpg")).toBe(true);
      expect(isValidImageType("image/png")).toBe(true);
      expect(isValidImageType("image/gif")).toBe(true);
      expect(isValidImageType("image/webp")).toBe(true);
      expect(isValidImageType("image/bmp")).toBe(true);
      expect(isValidImageType("image/svg+xml")).toBe(true);
    });

    it("should reject invalid image types", () => {
      expect(isValidImageType("text/plain")).toBe(false);
      expect(isValidImageType("application/pdf")).toBe(false);
      expect(isValidImageType("video/mp4")).toBe(false);
      expect(isValidImageType("audio/mp3")).toBe(false);
    });

    it("should handle case insensitive validation", () => {
      expect(isValidImageType("IMAGE/JPEG")).toBe(true);
      expect(isValidImageType("Image/PNG")).toBe(true);
      expect(isValidImageType("IMAGE/GIF")).toBe(true);
    });
  });

  describe("getFileExtensionFromMimeType", () => {
    it("should return correct extensions for image types", () => {
      expect(getFileExtensionFromMimeType("image/jpeg")).toBe("jpg");
      expect(getFileExtensionFromMimeType("image/jpg")).toBe("jpg");
      expect(getFileExtensionFromMimeType("image/png")).toBe("png");
      expect(getFileExtensionFromMimeType("image/gif")).toBe("gif");
      expect(getFileExtensionFromMimeType("image/webp")).toBe("webp");
      expect(getFileExtensionFromMimeType("image/bmp")).toBe("bmp");
      expect(getFileExtensionFromMimeType("image/svg+xml")).toBe("svg");
    });

    it("should return default extension for unknown types", () => {
      expect(getFileExtensionFromMimeType("image/unknown")).toBe("img");
      expect(getFileExtensionFromMimeType("application/pdf")).toBe("img");
    });

    it("should handle case insensitive mime types", () => {
      expect(getFileExtensionFromMimeType("IMAGE/JPEG")).toBe("jpg");
      expect(getFileExtensionFromMimeType("Image/PNG")).toBe("png");
    });
  });

  describe("imageUrlToBase64", () => {
    it("should convert image URL to base64", async () => {
      const mockArrayBuffer = new ArrayBuffer(4);
      const mockResponse = {
        ok: true,
        headers: {
          get: vi.fn().mockReturnValue("image/jpeg"),
        },
        arrayBuffer: vi.fn().mockResolvedValue(mockArrayBuffer),
      };

      mockFetch.mockResolvedValue(mockResponse);

      const result = await imageUrlToBase64("https://example.com/image.jpg");

      expect(result).toEqual({
        data: "AAAAAA==", // Base64 of empty 4-byte buffer
        contentType: "image/jpeg",
      });

      expect(mockFetch).toHaveBeenCalledWith("https://example.com/image.jpg");
    });

    it("should handle fetch errors", async () => {
      const mockResponse = {
        ok: false,
        statusText: "Not Found",
      };

      mockFetch.mockResolvedValue(mockResponse);

      const consoleSpy = vi.spyOn(console, "error").mockImplementation(() => {
        // Mock implementation - do nothing
      });

      const result = await imageUrlToBase64("https://example.com/nonexistent.jpg");

      expect(result).toBeNull();
      expect(consoleSpy).toHaveBeenCalledWith(
        "Error converting image URL to base64:",
        expect.any(Error)
      );

      consoleSpy.mockRestore();
    });

    it("should reject unsupported content types", async () => {
      const mockResponse = {
        ok: true,
        headers: {
          get: vi.fn().mockReturnValue("text/plain"),
        },
      };

      mockFetch.mockResolvedValue(mockResponse);

      const consoleSpy = vi.spyOn(console, "error").mockImplementation(() => {
        // Mock implementation - do nothing
      });

      const result = await imageUrlToBase64("https://example.com/text.txt");

      expect(result).toBeNull();
      expect(consoleSpy).toHaveBeenCalledWith(
        "Error converting image URL to base64:",
        expect.any(Error)
      );

      consoleSpy.mockRestore();
    });

    it("should use default content type when header is missing", async () => {
      const mockArrayBuffer = new ArrayBuffer(4);
      const mockResponse = {
        ok: true,
        headers: {
          get: vi.fn().mockReturnValue(null),
        },
        arrayBuffer: vi.fn().mockResolvedValue(mockArrayBuffer),
      };

      mockFetch.mockResolvedValue(mockResponse);

      const result = await imageUrlToBase64("https://example.com/image");

      expect(result).toEqual({
        data: "AAAAAA==",
        contentType: "image/jpeg", // Default fallback
      });
    });

    it("should handle network errors", async () => {
      mockFetch.mockRejectedValue(new Error("Network error"));

      const consoleSpy = vi.spyOn(console, "error").mockImplementation(() => {
        // Mock implementation - do nothing
      });

      const result = await imageUrlToBase64("https://example.com/image.jpg");

      expect(result).toBeNull();
      expect(consoleSpy).toHaveBeenCalledWith(
        "Error converting image URL to base64:",
        expect.any(Error)
      );

      consoleSpy.mockRestore();
    });
  });
});
