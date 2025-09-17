#!/usr/bin/env node

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import * as azdev from "azure-devops-node-api";
import { AccessToken, DefaultAzureCredential } from "@azure/identity";
import { configurePrompts } from "./prompts.js";
import { configureAllTools } from "./tools.js";
import { userAgent } from "./utils.js";
import { packageVersion } from "./version.js";
const args = process.argv.slice(2);
if (args.length === 0) {
  console.error("Usage: mcp-server-azuredevops <organization_name>");
  process.exit(1);
}

export const orgName = args[0];
const orgUrl = "https://dev.azure.com/" + orgName;

async function getAzureDevOpsToken(): Promise<AccessToken> {
<<<<<<< Updated upstream
  process.env.AZURE_TOKEN_CREDENTIALS = "dev";
  const credential = new DefaultAzureCredential(); // CodeQL [SM05138] resolved by explicitly setting AZURE_TOKEN_CREDENTIALS
=======
  // If PAT is available, we don't need to get an Azure token, but we still need to return an AccessToken-like object
  if (process.env.AZURE_DEVOPS_PAT) {
    // Return a mock AccessToken since we'll use Basic auth instead
    return {
      token: process.env.AZURE_DEVOPS_PAT,
      expiresOnTimestamp: Date.now() + 3600000 // 1 hour from now
    };
  }

  if (process.env.ADO_MCP_AZURE_TOKEN_CREDENTIALS) {
    process.env.AZURE_TOKEN_CREDENTIALS = process.env.ADO_MCP_AZURE_TOKEN_CREDENTIALS;
  } else {
    process.env.AZURE_TOKEN_CREDENTIALS = "dev";
  }
  let credential: TokenCredential = new DefaultAzureCredential(); // CodeQL [SM05138] resolved by explicitly setting AZURE_TOKEN_CREDENTIALS
  if (tenantId) {
    // Use Azure CLI credential if tenantId is provided for multi-tenant scenarios
    const azureCliCredential = new AzureCliCredential({ tenantId });
    credential = new ChainedTokenCredential(azureCliCredential, credential);
  }

>>>>>>> Stashed changes
  const token = await credential.getToken("499b84ac-1321-427f-aa17-267ca6975798/.default");
  return token;
}

<<<<<<< Updated upstream
async function getAzureDevOpsClient(): Promise<azdev.WebApi> {
  const token = await getAzureDevOpsToken();
  const authHandler = azdev.getBearerHandler(token.token);
  const connection = new azdev.WebApi(orgUrl, authHandler, undefined, {
    productName: "AzureDevOps.MCP",
    productVersion: packageVersion,
    userAgent: userAgent,
  });
  return connection;
=======
function getAzureDevOpsClient(userAgentComposer: UserAgentComposer): () => Promise<azdev.WebApi> {
  return async () => {
    // Use PAT authentication if available
    if (process.env.AZURE_DEVOPS_PAT) {
      const authHandler = azdev.getPersonalAccessTokenHandler(process.env.AZURE_DEVOPS_PAT);
      const connection = new azdev.WebApi(orgUrl, authHandler, undefined, {
        productName: "AzureDevOps.MCP",
        productVersion: packageVersion,
        userAgent: userAgentComposer.userAgent,
      });
      return connection;
    }

    // Fall back to Azure authentication
    const token = await getAzureDevOpsToken();
    const authHandler = azdev.getBearerHandler(token.token);
    const connection = new azdev.WebApi(orgUrl, authHandler, undefined, {
      productName: "AzureDevOps.MCP",
      productVersion: packageVersion,
      userAgent: userAgentComposer.userAgent,
    });
    return connection;
  };
>>>>>>> Stashed changes
}

async function main() {
  const server = new McpServer({
    name: "Azure DevOps MCP Server",
    version: packageVersion,
  });

  configurePrompts(server);

  configureAllTools(server, getAzureDevOpsToken, getAzureDevOpsClient);

  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  process.exit(1);
});
