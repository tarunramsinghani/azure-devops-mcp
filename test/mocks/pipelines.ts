// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export const mockBuildDefinitions = [
  {
    id: 1,
    name: "CI Build",
    project: { id: "proj1", name: "Test Project" },
    repository: { id: "repo1", name: "Test Repo" },
    path: "\\CI",
  },
  {
    id: 2,
    name: "Release Build",
    project: { id: "proj1", name: "Test Project" },
    repository: { id: "repo1", name: "Test Repo" },
    path: "\\Release",
  },
];

export const mockBuilds = [
  {
    id: 123,
    buildNumber: "20250107.1",
    status: "completed",
    result: "succeeded",
    project: { id: "proj1", name: "Test Project" },
    definition: { id: 1, name: "CI Build" },
  },
  {
    id: 124,
    buildNumber: "20250107.2",
    status: "inProgress",
    result: null,
    project: { id: "proj1", name: "Test Project" },
    definition: { id: 1, name: "CI Build" },
  },
];

export const mockBuildLogs = [
  {
    id: 1,
    type: "Container",
    url: "https://dev.azure.com/org/proj/_apis/build/builds/123/logs/1",
  },
  {
    id: 2,
    type: "Task",
    url: "https://dev.azure.com/org/proj/_apis/build/builds/123/logs/2",
  },
];

export const mockBuildLogContent = `
##[section]Starting: Build
Task         : Command line
Description  : Run a command line script using Bash on Linux and macOS and cmd.exe on Windows
Version      : 2.164.1
Author       : Microsoft Corporation
Help         : https://docs.microsoft.com/azure/devops/pipelines/tasks/utility/command-line

##[command]echo "Hello World"
Hello World
##[section]Finishing: Build
`;

export const mockBuildChanges = [
  {
    id: "commit123",
    message: "Fix bug in authentication",
    author: { displayName: "John Doe", uniqueName: "john@example.com" },
    timestamp: "2025-01-07T10:00:00Z",
  },
  {
    id: "commit456",
    message: "Add new feature",
    author: { displayName: "Jane Smith", uniqueName: "jane@example.com" },
    timestamp: "2025-01-07T09:30:00Z",
  },
];

export const mockBuildReport = {
  buildId: 123,
  content: "Build completed successfully",
  type: "BuildSummary",
};

export const mockUpdateBuildStageResponse = {
  state: 0,
  name: "Build",
  forceRetryAllJobs: false,
};

export const mockBuildTimeline = {
  records: [
    {
      id: "record1",
      name: "Build Stage",
      type: "Stage",
      state: 2, // Completed
      result: 0, // Succeeded
      startTime: "2025-01-07T10:00:00Z",
      finishTime: "2025-01-07T10:05:00Z",
      log: { id: 1 },
      parentId: null,
      order: 1,
    },
    {
      id: "record2",
      name: "NuGet restore",
      type: "Task",
      state: 2, // Completed
      result: 0, // Succeeded
      startTime: "2025-01-07T10:01:00Z",
      finishTime: "2025-01-07T10:02:00Z",
      log: { id: 2 },
      parentId: "record1",
      order: 2,
    },
    {
      id: "record3",
      name: "Run Tests",
      type: "Task",
      state: 2, // Completed
      result: 2, // Failed
      startTime: "2025-01-07T10:03:00Z",
      finishTime: "2025-01-07T10:04:00Z",
      log: { id: 3 },
      parentId: "record1",
      order: 3,
    },
  ],
};
