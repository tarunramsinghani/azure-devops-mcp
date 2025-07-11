import { AccessToken } from "@azure/identity";
import { describe, expect, it, beforeEach } from "@jest/globals";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import { TaskResult, TimelineRecordState } from "azure-devops-node-api/interfaces/BuildInterfaces.js";
import { StageUpdateType } from "azure-devops-node-api/interfaces/BuildInterfaces.js";
import { configureBuildTools } from "../../../src/tools/builds";
import { apiVersion } from "../../../src/utils.js";
import { mockUpdateBuildStageResponse, mockBuildLogs, mockBuildTimeline } from "../../mocks/builds";

// Mock fetch globally
global.fetch = jest.fn() as jest.MockedFunction<typeof fetch>;

type TokenProviderMock = () => Promise<AccessToken>;
type ConnectionProviderMock = () => Promise<WebApi>;

describe("configureBuildTools", () => {
  let server: McpServer;
  let tokenProvider: TokenProviderMock;
  let connectionProvider: ConnectionProviderMock;
  let mockConnection: { getBuildApi: jest.Mock; serverUrl: string };

  beforeEach(() => {
    server = { tool: jest.fn() } as unknown as McpServer;
    tokenProvider = jest.fn();
    mockConnection = {
      getBuildApi: jest.fn(),
      serverUrl: "https://dev.azure.com/test-org",
    };
    connectionProvider = jest.fn().mockResolvedValue(mockConnection);
    (global.fetch as jest.MockedFunction<typeof fetch>).mockClear();
  });

  describe("tool registration", () => {
    it("registers build tools on the server", () => {
      configureBuildTools(server, tokenProvider, connectionProvider);
      expect(server.tool as jest.Mock).toHaveBeenCalled();
    });
  });

  describe("update_build_stage tool", () => {
    it("should update build stage with correct parameters and return the expected result", async () => {
      configureBuildTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "build_update_build_stage");
      if (!call) throw new Error("build_update_build_stage tool not registered");
      const [, , , handler] = call;

      // Mock the token provider
      const mockToken = { token: "mock-token" };
      (tokenProvider as jest.Mock).mockResolvedValue(mockToken);

      // Mock successful fetch response
      const mockResponse = {
        ok: true,
        text: jest.fn().mockResolvedValue(JSON.stringify(mockUpdateBuildStageResponse)),
      };
      (global.fetch as jest.MockedFunction<typeof fetch>).mockResolvedValue(mockResponse as unknown as Response);

      const params = {
        project: "test-project",
        buildId: 123,
        stageName: "Build",
        status: StageUpdateType.Retry,
        forceRetryAllJobs: true,
      };

      const result = await handler(params);

      expect(global.fetch).toHaveBeenCalledWith(`https://dev.azure.com/test-org/test-project/_apis/build/builds/123/stages/Build?api-version=${apiVersion}`, {
        method: "PATCH",
        headers: {
          "Content-Type": "application/json",
          "Authorization": "Bearer mock-token",
        },
        body: JSON.stringify({
          forceRetryAllJobs: true,
          state: StageUpdateType.Retry.valueOf(),
        }),
      });
      expect(result.content[0].text).toBe(JSON.stringify(JSON.stringify(mockUpdateBuildStageResponse), null, 2));
      expect(result.isError).toBeUndefined();
    });

    it("should handle HTTP errors correctly", async () => {
      configureBuildTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "build_update_build_stage");
      if (!call) throw new Error("build_update_build_stage tool not registered");
      const [, , , handler] = call;

      // Mock the token provider
      const mockToken = { token: "mock-token" };
      (tokenProvider as jest.Mock).mockResolvedValue(mockToken);

      // Mock failed fetch response
      const mockResponse = {
        ok: false,
        status: 404,
        text: jest.fn().mockResolvedValue("Build stage not found"),
      };
      (global.fetch as jest.MockedFunction<typeof fetch>).mockResolvedValue(mockResponse as unknown as Response);

      const params = {
        project: "test-project",
        buildId: 999,
        stageName: "NonExistentStage",
        status: StageUpdateType.Retry,
        forceRetryAllJobs: false,
      };

      await expect(handler(params)).rejects.toThrow("Failed to update build stage: 404 Build stage not found");

      expect(global.fetch).toHaveBeenCalledWith(`https://dev.azure.com/test-org/test-project/_apis/build/builds/999/stages/NonExistentStage?api-version=${apiVersion}`, {
        method: "PATCH",
        headers: {
          "Content-Type": "application/json",
          "Authorization": "Bearer mock-token",
        },
        body: JSON.stringify({
          forceRetryAllJobs: false,
          state: StageUpdateType.Retry.valueOf(),
        }),
      });
    });

    it("should handle network errors correctly", async () => {
      configureBuildTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "build_update_build_stage");
      if (!call) throw new Error("build_update_build_stage tool not registered");
      const [, , , handler] = call;

      // Mock the token provider
      const mockToken = { token: "mock-token" };
      (tokenProvider as jest.Mock).mockResolvedValue(mockToken);

      // Mock network error
      const networkError = new Error("Network connection failed");
      (global.fetch as jest.MockedFunction<typeof fetch>).mockRejectedValue(networkError);

      const params = {
        project: "test-project",
        buildId: 123,
        stageName: "Build",
        status: StageUpdateType.Retry,
        forceRetryAllJobs: false,
      };

      await expect(handler(params)).rejects.toThrow("Network connection failed");

      expect(global.fetch).toHaveBeenCalledWith(`https://dev.azure.com/test-org/test-project/_apis/build/builds/123/stages/Build?api-version=${apiVersion}`, {
        method: "PATCH",
        headers: {
          "Content-Type": "application/json",
          "Authorization": "Bearer mock-token",
        },
        body: JSON.stringify({
          forceRetryAllJobs: false,
          state: StageUpdateType.Retry.valueOf(),
        }),
      });
    });

    it("should handle token provider errors correctly", async () => {
      configureBuildTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "build_update_build_stage");
      if (!call) throw new Error("build_update_build_stage tool not registered");
      const [, , , handler] = call;

      // Mock token provider error
      const tokenError = new Error("Failed to get access token");
      (tokenProvider as jest.Mock).mockRejectedValue(tokenError);

      const params = {
        project: "test-project",
        buildId: 123,
        stageName: "Build",
        status: StageUpdateType.Retry,
        forceRetryAllJobs: false,
      };

      await expect(handler(params)).rejects.toThrow("Failed to get access token");

      // Should not call fetch if token provider fails
      expect(global.fetch).not.toHaveBeenCalled();
    });

    it("should handle different StageUpdateType values correctly", async () => {
      configureBuildTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "build_update_build_stage");
      if (!call) throw new Error("build_update_build_stage tool not registered");
      const [, , , handler] = call;

      const mockToken = { token: "mock-token" };
      (tokenProvider as jest.Mock).mockResolvedValue(mockToken);

      const mockResponse = {
        ok: true,
        text: jest.fn().mockResolvedValue(JSON.stringify(mockUpdateBuildStageResponse)),
      };
      (global.fetch as jest.MockedFunction<typeof fetch>).mockResolvedValue(mockResponse as unknown as Response);

      const params = {
        project: "test-project",
        buildId: 123,
        stageName: "Deploy",
        status: StageUpdateType.Cancel,
        forceRetryAllJobs: false,
      };

      await handler(params);

      expect(global.fetch).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({
          body: JSON.stringify({
            forceRetryAllJobs: false,
            state: StageUpdateType.Cancel.valueOf(),
          }),
        })
      );
    });
  });

  describe("get_log tool", () => {
    it("should retrieve logs with step information and status", async () => {
      configureBuildTools(server, tokenProvider, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "build_get_log");
      if (!call) throw new Error("build_get_log tool not registered");
      const [, , , handler] = call;

      // Mock the build API methods
      const mockBuildApi = {
        getBuildLogs: jest.fn().mockResolvedValue(mockBuildLogs),
        getBuildTimeline: jest.fn().mockResolvedValue(mockBuildTimeline)
      };
      mockConnection.getBuildApi.mockResolvedValue(mockBuildApi);

      const params = {
        project: "test-project",
        buildId: 123
      };

      const result = await handler(params);
      const response = JSON.parse(result.content[0].text);

      // Verify API calls
      expect(mockBuildApi.getBuildLogs).toHaveBeenCalledWith("test-project", 123);
      expect(mockBuildApi.getBuildTimeline).toHaveBeenCalledWith("test-project", 123);

      // Verify response structure
      expect(response).toHaveProperty("logs");
      expect(response).toHaveProperty("steps");
      expect(response).toHaveProperty("summary");
      expect(response).toHaveProperty("buildId", 123);

      // Verify enhanced logs with step information
      expect(response.logs).toHaveLength(2);
      expect(response.logs[0]).toHaveProperty("stepInfo");
      expect(response.logs[0].stepInfo).toHaveProperty("stepName", "Build Stage");
      expect(response.logs[0].stepInfo).toHaveProperty("state", TimelineRecordState.Completed);
      expect(response.logs[0].stepInfo).toHaveProperty("result", TaskResult.Succeeded);

      // Verify steps information
      expect(response.steps).toHaveLength(3); // 1 Stage + 2 Tasks
      expect(response.steps[0]).toMatchObject({
        name: "Build Stage",
        type: "Stage",
        state: TimelineRecordState.Completed,
        result: TaskResult.Succeeded
      });

      // Verify summary
      expect(response.summary).toMatchObject({
        totalLogs: 2,
        totalSteps: 3,
        passedSteps: 2,
        failedSteps: 1,
        skippedSteps: 0,
        inProgressSteps: 0
      });
    });
  });
});
