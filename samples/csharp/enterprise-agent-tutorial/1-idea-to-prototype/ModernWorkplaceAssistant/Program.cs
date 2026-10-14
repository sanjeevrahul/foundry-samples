using System;
using System.Collections.Generic;
using System.Threading.Tasks;

using Azure.AI.Projects;
using Azure.AI.Projects.Agents;
using Azure.AI.Extensions.OpenAI;
using Azure.Identity;

using DotNetEnv;
using OpenAI.Responses;

namespace ModernWorkplaceAssistant;

internal class Program
{
    private const string AgentName = "Modern Workplace Assistant";

    private static async Task Main(string[] args)
    {
        Console.WriteLine();
        Console.WriteLine("======================================================================");
        Console.WriteLine("        MICROSOFT FOUNDRY - MODERN WORKPLACE ASSISTANT");
        Console.WriteLine("======================================================================");
        Console.WriteLine();

        // ---------------------------------------------------------------
        // Load environment variables
        // ---------------------------------------------------------------

        Env.Load();

        string projectEndpoint =
            Environment.GetEnvironmentVariable(
                "FOUNDRY_PROJECT_ENDPOINT")
            ?? throw new InvalidOperationException(
                "FOUNDRY_PROJECT_ENDPOINT is not configured.");

        string modelDeploymentName =
            Environment.GetEnvironmentVariable(
                "FOUNDRY_MODEL_NAME")
            ?? throw new InvalidOperationException(
                "FOUNDRY_MODEL_NAME is not configured.");

        string? sharePointConnectionName =
            Environment.GetEnvironmentVariable(
                "SHAREPOINT_CONNECTION_NAME");

        string? mcpServerUrl =
            Environment.GetEnvironmentVariable(
                "MCP_SERVER_URL");

        Console.WriteLine(
            $"Project Endpoint : {projectEndpoint}");

        Console.WriteLine(
            $"Model Deployment : {modelDeploymentName}");

        Console.WriteLine();

        // ---------------------------------------------------------------
        // Authenticate
        // ---------------------------------------------------------------

        DefaultAzureCredential credential = new();

        AIProjectClient projectClient =
            new(
                endpoint: new Uri(projectEndpoint),
                tokenProvider: credential);

        Console.WriteLine(
            "Connected to Microsoft Foundry.");
        Console.WriteLine();

        // ---------------------------------------------------------------
        // Create agent tools
        // ---------------------------------------------------------------

        List<ProjectsAgentTool> tools = new();

        // ===============================================================
        // SharePoint
        // ===============================================================

        if (!string.IsNullOrWhiteSpace(
                sharePointConnectionName))
        {
            Console.WriteLine(
                "Configuring SharePoint...");

            try
            {
                AIProjectConnection connection =
                    await projectClient
                        .Connections
                        .GetConnectionAsync(
                            sharePointConnectionName,
                            includeCredentials: false);

                Console.WriteLine(
                    $"SharePoint connection: {connection.Id}");

                SharePointGroundingToolOptions
                    sharePointOptions = new();

                sharePointOptions
                    .ProjectConnections
                    .Add(
                        new ToolProjectConnection(
                            projectConnectionId:
                                connection.Id));

                SharepointPreviewTool sharePointTool =
                    new(
                        sharePointOptions);

                /*
                 * SharepointPreviewTool derives from
                 * ProjectsAgentTool, so it can be added directly.
                 */
                tools.Add(sharePointTool);

                Console.WriteLine(
                    "SharePoint tool configured.");
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    $"SharePoint configuration failed: {ex.Message}");

                Console.WriteLine(
                    "Continuing without SharePoint.");
            }
        }
        else
        {
            Console.WriteLine(
                "SHAREPOINT_CONNECTION_NAME not configured.");

            Console.WriteLine(
                "SharePoint will not be enabled.");
        }

        Console.WriteLine();

        // ===============================================================
        // Microsoft Learn MCP
        // ===============================================================

        /*
         * IMPORTANT:
         *
         * In your installed 2.0.0 package, OpenAI.Responses.McpTool
         * cannot be directly inserted into List<ProjectsAgentTool>.
         *
         * Therefore we don't add the OpenAI MCP object to the
         * ProjectsAgentTool collection here.
         *
         * MCP is handled through the Responses API instead.
         */

        McpTool? mcpTool = null;

        if (!string.IsNullOrWhiteSpace(mcpServerUrl))
        {
            Console.WriteLine(
                "Configuring Microsoft Learn MCP...");

            Console.WriteLine(
                $"MCP Server: {mcpServerUrl}");

            try
            {
                mcpTool =
                    ResponseTool.CreateMcpTool(
                        serverLabel:
                            "Microsoft_Learn_Documentation",
                        serverUri:
                            new Uri(mcpServerUrl),
                        toolCallApprovalPolicy:
                            new McpToolCallApprovalPolicy(
                                GlobalMcpToolCallApprovalPolicy
                                    .AlwaysRequireApproval));

                Console.WriteLine(
                    "Microsoft Learn MCP configured.");
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    $"MCP configuration failed: {ex.Message}");

                mcpTool = null;
            }
        }
        else
        {
            Console.WriteLine(
                "MCP_SERVER_URL not configured.");

            Console.WriteLine(
                "Microsoft Learn MCP will not be enabled.");
        }

        Console.WriteLine();

        // ---------------------------------------------------------------
        // Agent instructions
        // ---------------------------------------------------------------

        string instructions = """
            You are the Modern Workplace Assistant.

            You help enterprise employees find reliable information
            and answer workplace questions.

            Use SharePoint for:

            - Internal company policies
            - Internal procedures
            - Company documentation
            - Organisation-specific information
            - Internal processes

            Use Microsoft Learn for:

            - Microsoft Azure
            - Microsoft Foundry
            - Microsoft 365
            - Power Platform
            - .NET
            - Microsoft development technologies

            When information from multiple sources is required,
            combine the information carefully.

            Important rules:

            - Never invent company policies.
            - Prefer authoritative internal information for
              organisation-specific questions.
            - Prefer Microsoft Learn for Microsoft technical questions.
            - Clearly distinguish internal information from
              general Microsoft guidance.
            - If information cannot be found, say so.
            - Do not make unsupported assumptions.
            - Give concise and practical answers.
            """;

        // ---------------------------------------------------------------
        // Create agent definition
        // ---------------------------------------------------------------

        Console.WriteLine(
            "Creating Foundry agent...");

        DeclarativeAgentDefinition agentDefinition =
            new(
                model: modelDeploymentName)
            {
                Instructions = instructions
            };

        foreach (ProjectsAgentTool tool in tools)
        {
            agentDefinition.Tools.Add(tool);
        }

        // ---------------------------------------------------------------
        // Create agent version
        // ---------------------------------------------------------------

        ProjectsAgentVersion agentVersion =
            await projectClient
                .AgentAdministrationClient
                .CreateAgentVersionAsync(
                    agentName: AgentName,
                    options:
                        new(agentDefinition));

        Console.WriteLine();
        Console.WriteLine(
            "Agent created successfully.");

        Console.WriteLine(
            $"Name    : {agentVersion.Name}");

        Console.WriteLine(
            $"Version : {agentVersion.Version}");

        Console.WriteLine();

        // ---------------------------------------------------------------
        // Create project Responses client
        // ---------------------------------------------------------------

        AgentReference agentReference =
            new(
                name: agentVersion.Name,
                version: agentVersion.Version);

        ProjectResponsesClient responsesClient =
            projectClient
                .ProjectOpenAIClient
                .GetProjectResponsesClientForAgent(
                    agentReference);

        Console.WriteLine(
            "Responses client created.");

        Console.WriteLine();

        // ---------------------------------------------------------------
        // Interactive assistant
        // ---------------------------------------------------------------

        Console.WriteLine(
            "======================================================================");
        Console.WriteLine(
            "                    MODERN WORKPLACE ASSISTANT");
        Console.WriteLine(
            "======================================================================");

        Console.WriteLine();

        Console.WriteLine(
            "Ask a question.");

        Console.WriteLine(
            "Type 'exit' to quit.");

        Console.WriteLine();

        // ---------------------------------------------------------------
        // Conversation loop
        // ---------------------------------------------------------------

        while (true)
        {
            Console.Write("You: ");

            string? userInput =
                Console.ReadLine();

            if (string.IsNullOrWhiteSpace(
                    userInput))
            {
                continue;
            }

            if (string.Equals(
                    userInput,
                    "exit",
                    StringComparison.OrdinalIgnoreCase))
            {
                break;
            }

            try
            {
                Console.WriteLine();
                Console.WriteLine(
                    "Agent is thinking...");

                ResponseResult response =
                    await RunAgentAsync(
                        responsesClient,
                        userInput,
                        mcpTool);

                Console.WriteLine();
                Console.WriteLine(
                    "Assistant:");

                Console.WriteLine(
                    response.GetOutputText());

                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine(
                    $"ERROR: {ex.Message}");

                Console.WriteLine();
            }
        }

        // ---------------------------------------------------------------
        // Cleanup
        // ---------------------------------------------------------------

        Console.WriteLine();
        Console.WriteLine(
            "Cleaning up agent...");

        try
        {
            await projectClient
                .AgentAdministrationClient
                .DeleteAgentVersionAsync(
                    agentName:
                        agentVersion.Name,
                    agentVersion:
                        agentVersion.Version);

            Console.WriteLine(
                "Agent deleted.");
        }
        catch (Exception ex)
        {
            Console.WriteLine(
                $"Could not delete agent: {ex.Message}");
        }

        Console.WriteLine();
        Console.WriteLine(
            "Goodbye!");
    }

    // ===================================================================
    // Run agent
    // ===================================================================

    private static async Task<ResponseResult> RunAgentAsync(
        ProjectResponsesClient responsesClient,
        string userInput,
        McpTool? mcpTool)
    {
        CreateResponseOptions options =
            new();

        options.InputItems.Add(
            ResponseItem.CreateUserMessageItem(
                userInput));

        /*
         * MCP is attached to the Responses request,
         * not to ProjectsAgentTool.
         */
        if (mcpTool is not null)
        {
            options.Tools.Add(mcpTool);
        }

        ResponseResult? response = null;

        while (true)
        {
            response =
                await responsesClient
                    .CreateResponseAsync(
                        options);

            List<ResponseItem>
                approvalResponses = new();

            foreach (ResponseItem item
                     in response.OutputItems)
            {
                // -------------------------------------------------------
                // MCP approval
                // -------------------------------------------------------

                if (item
                    is McpToolCallApprovalRequestItem approval)
                {
                    Console.WriteLine();
                    Console.WriteLine(
                        "MCP approval requested.");

                    Console.WriteLine(
                        $"Server: {approval.ServerLabel}");

                    bool approved =
                        string.Equals(
                            approval.ServerLabel,
                            "Microsoft_Learn_Documentation",
                            StringComparison.OrdinalIgnoreCase);

                    Console.WriteLine(
                        approved
                            ? "MCP call approved."
                            : "MCP call rejected.");

                    approvalResponses.Add(
                        ResponseItem
                            .CreateMcpApprovalResponseItem(
                                approvalRequestId:
                                    approval.Id,
                                approved:
                                    approved));
                }
            }

            // -----------------------------------------------------------
            // No approval required
            // -----------------------------------------------------------

            if (approvalResponses.Count == 0)
            {
                break;
            }

            // -----------------------------------------------------------
            // Continue previous response
            // -----------------------------------------------------------

            options =
                new CreateResponseOptions
                {
                    PreviousResponseId =
                        response.Id
                };

            foreach (ResponseItem approvalResponse
                     in approvalResponses)
            {
                options.InputItems.Add(
                    approvalResponse);
            }

            /*
             * Preserve MCP tool for the continuation request.
             */
            if (mcpTool is not null)
            {
                options.Tools.Add(mcpTool);
            }
        }

        return response
            ?? throw new InvalidOperationException(
                "The agent did not return a response.");
    }
}