// <imports_and_includes>
using Azure;
using Azure.AI.Projects;
using Azure.AI.Projects.Agents;
using Azure.AI.Extensions.OpenAI;
using Azure.Core;
using Azure.Identity;
using DotNetEnv;
using OpenAI.Responses;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Threading.Tasks;
// </imports_and_includes>

#pragma warning disable OPENAI001

class EvaluateProgram
{
    private static string GetFile(
        string name,
        [CallerFilePath] string pth = "")
    {
        var dirName = Path.GetDirectoryName(pth) ?? "";
        return Path.Combine(dirName, "..", "shared", name);
    }

    static async Task Main(string[] args)
    {
        // ============================================================
        // Load environment variables
        // ============================================================

        Env.Load(GetFile(".env"));

        var projectEndpoint =
            Environment.GetEnvironmentVariable(
                "FOUNDRY_PROJECT_ENDPOINT");

        var modelDeploymentName =
            Environment.GetEnvironmentVariable(
                "FOUNDRY_MODEL_NAME");

        var sharepointConnectionId =
            Environment.GetEnvironmentVariable(
                "SHAREPOINT_CONNECTION_ID");

        var mcpServerUrl =
            Environment.GetEnvironmentVariable(
                "MCP_SERVER_URL");

        var tenantId =
            Environment.GetEnvironmentVariable(
                "AI_FOUNDRY_TENANT_ID");

        if (string.IsNullOrWhiteSpace(projectEndpoint))
        {
            throw new InvalidOperationException(
                "FOUNDRY_PROJECT_ENDPOINT is not configured.");
        }

        if (string.IsNullOrWhiteSpace(modelDeploymentName))
        {
            throw new InvalidOperationException(
                "FOUNDRY_MODEL_NAME is not configured.");
        }

        // ============================================================
        // Authentication
        // ============================================================

        TokenCredential credential;

        if (!string.IsNullOrEmpty(tenantId))
        {
            credential =
                new AzureCliCredential(
                    new AzureCliCredentialOptions
                    {
                        TenantId = tenantId
                    });
        }
        else
        {
            credential =
                new DefaultAzureCredential();
        }

        // ============================================================
        // Create Foundry project client
        // ============================================================

        AIProjectClient client =
            new(
                new Uri(projectEndpoint),
                credential);

        Console.WriteLine(
            "🧪 Modern Workplace Assistant Evaluation\n");

        // ============================================================
        // Agent instructions
        // ============================================================

        var instructions = """
            You are a Modern Workplace Assistant for Contoso Corporation.

            Answer questions using available tools and provide specific,
            detailed responses.

            Use SharePoint for internal company information and
            Microsoft Learn for Microsoft technical information.

            Do not invent information.
            """;

        // ============================================================
        // Create agent definition
        // ============================================================

        DeclarativeAgentDefinition agentDefinition =
            new(
                model: modelDeploymentName)
            {
                Instructions = instructions
            };

        // ============================================================
        // SharePoint tool
        // ============================================================

        if (!string.IsNullOrEmpty(
                sharepointConnectionId))
        {
            try
            {
                SharePointGroundingToolOptions
                    sharepointToolOption = new()
                    {
                        ProjectConnections =
                        {
                            new ToolProjectConnection(
                                projectConnectionId:
                                    sharepointConnectionId)
                        }
                    };

                SharepointPreviewTool sharepointTool =
                    new(
                        sharepointToolOption);

                /*
                 * SharepointPreviewTool derives from
                 * ProjectsAgentTool and can therefore be
                 * registered on the agent definition.
                 */
                agentDefinition.Tools.Add(
                    sharepointTool);

                Console.WriteLine(
                    "✅ SharePoint configured for evaluation");
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    $"⚠️ SharePoint unavailable: {ex.Message}");
            }
        }

        // ============================================================
        // Microsoft Learn MCP
        // ============================================================

        McpTool? mcpTool = null;

        if (!string.IsNullOrEmpty(
                mcpServerUrl))
        {
            try
            {
                mcpTool =
                    ResponseTool.CreateMcpTool(
                        serverLabel:
                            "microsoft_learn",

                        serverUri:
                            new Uri(mcpServerUrl),

                        toolCallApprovalPolicy:
                            new McpToolCallApprovalPolicy(
                                GlobalMcpToolCallApprovalPolicy
                                    .NeverRequireApproval));

                /*
                 * IMPORTANT:
                 *
                 * With the installed Azure.AI.Extensions.OpenAI
                 * 2.0.0 / Azure.AI.Projects.Agents 2.0.0 combination,
                 * OpenAI.Responses.McpTool cannot be directly added
                 * to DeclarativeAgentDefinition.Tools.
                 *
                 * Therefore MCP is attached to the Responses request
                 * later.
                 */

                Console.WriteLine(
                    "✅ MCP configured for evaluation");
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    $"⚠️ MCP unavailable: {ex.Message}");
            }
        }

        Console.WriteLine();

        // ============================================================
        // Create agent version
        // ============================================================

        ProjectsAgentVersion agent =
            await client
                .AgentAdministrationClient
                .CreateAgentVersionAsync(
                    agentName:
                        "EvaluationAgent",

                    options:
                        new(agentDefinition));

        Console.WriteLine(
            $"✅ Evaluation agent created: " +
            $"{agent.Name} / {agent.Version}");

        Console.WriteLine();

        // ============================================================
        // Load test data
        // ============================================================

        var questions =
            File.ReadAllLines(
                GetFile("questions.jsonl"))
            .Select(
                line =>
                    JsonSerializer.Deserialize<JsonElement>(
                        line))
            .ToList();

        // ============================================================
        // Create project Responses client
        // ============================================================

        AgentReference agentReference =
            new(
                name: agent.Name,
                version: agent.Version);

        ProjectResponsesClient responseClient =
            client
                .ProjectOpenAIClient
                .GetProjectResponsesClientForAgent(
                    agentReference);

        // ============================================================
        // Run batch evaluation
        // ============================================================

        var results =
            new List<object>();

        Console.WriteLine(
            $"Running {questions.Count} evaluation questions...\n");

        for (int i = 0;
             i < questions.Count;
             i++)
        {
            var q = questions[i];

            var question =
                q.GetProperty("question")
                 .GetString()!;

            string[] expectedKeywords =
                Array.Empty<string>();

            if (q.TryGetProperty(
                    "expected_keywords",
                    out var keywordsElem))
            {
                expectedKeywords =
                    keywordsElem
                        .EnumerateArray()
                        .Select(
                            e => e.GetString()!)
                        .ToArray();
            }

            Console.WriteLine(
                $"Question {i + 1}/{questions.Count}: " +
                $"{question}");

            string response = "";

            try
            {
                // ====================================================
                // Create response options
                // ====================================================

                CreateResponseOptions options =
                    new();

                options.InputItems.Add(
                    ResponseItem.CreateUserMessageItem(
                        question));

                // Add MCP to the response request
                if (mcpTool is not null)
                {
                    options.Tools.Add(
                        mcpTool);
                }

                // ====================================================
                // Execute agent
                // ====================================================

                ResponseResult openAIResponse =
                    await responseClient
                        .CreateResponseAsync(
                            options);

                response =
                    openAIResponse.GetOutputText();
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    $"   ⚠️ Error: {ex.Message}");

                response = "";
            }

            // ========================================================
            // Evaluation
            // ========================================================

            bool passed =
                response.Length > 50;

            if (expectedKeywords.Length > 0)
            {
                passed =
                    passed &&
                    expectedKeywords.Any(
                        k =>
                            response.Contains(
                                k,
                                StringComparison.OrdinalIgnoreCase));
            }

            Console.WriteLine(
                $"   Status: " +
                $"{(passed ? "✅ PASS" : "❌ FAIL")}");

            Console.WriteLine(
                $"   Response length: " +
                $"{response.Length} characters\n");

            results.Add(
                new
                {
                    question,
                    response,
                    passed,
                    response_length =
                        response.Length
                });
        }

        // ============================================================
        // Evaluation results
        // ============================================================

        var summary =
            new
            {
                total_questions =
                    questions.Count,

                passed =
                    results.Count(
                        r =>
                            ((dynamic)r).passed),

                failed =
                    results.Count(
                        r =>
                            !((dynamic)r).passed),

                results
            };

        var json =
            JsonSerializer.Serialize(
                summary,
                new JsonSerializerOptions
                {
                    WriteIndented = true
                });

        File.WriteAllText(
            "evaluation_results.json",
            json);

        // ============================================================
        // Print summary
        // ============================================================

        Console.WriteLine(
            "📊 Evaluation Complete:");

        Console.WriteLine(
            $"   Total: {summary.total_questions}");

        Console.WriteLine(
            $"   Passed: {summary.passed}");

        Console.WriteLine(
            $"   Failed: {summary.failed}");

        Console.WriteLine(
            "\n📄 Results saved to " +
            "evaluation_results.json");

        // ============================================================
        // Cleanup
        // ============================================================

        /*
         * Agents are versioned in SDK 2.0.
         *
         * Uncomment if you want to delete the evaluation agent
         * after execution.
         */

        /*
        await client
            .AgentAdministrationClient
            .DeleteAgentVersionAsync(
                agentName: agent.Name,
                agentVersion: agent.Version);

        Console.WriteLine(
            "🗑️ Evaluation agent deleted.");
        */
    }
}