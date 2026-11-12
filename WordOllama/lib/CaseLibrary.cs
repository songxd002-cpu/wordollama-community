using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib
{
    /// <summary>
    /// 案例库：由 Agent 面板的“调用案例库”勾选框控制，连接服务器端的案例检索 MCP 服务。
    /// 勾选框本身就是授权：未勾选时不连接、不注册任何案例库工具；
    /// 勾选时只注册白名单内的只读检索工具，与通用“外部工具”开关互不影响。
    /// </summary>
    public static class CaseLibrary
    {
        public const string ServerName = "case-library";
        public const string SearchToolName = "search_cases";
        public const string GetToolName = "get_case";

        private static readonly string[] AllowedToolNames = { SearchToolName, GetToolName };

        public static IReadOnlyList<string> ToolNames => AllowedToolNames;

        public static bool IsAllowedTool(string name) =>
            AllowedToolNames.Contains(name ?? "", StringComparer.Ordinal);

        /// <summary>
        /// 远程地址必须是 HTTPS；只有本机回环地址允许 HTTP，便于本地联调。空字符串视为未配置。
        /// </summary>
        public static bool IsValidEndpoint(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return true;
            if (!Uri.TryCreate(value.Trim(), UriKind.Absolute, out Uri uri)) return false;
            if (uri.Scheme == Uri.UriSchemeHttps) return true;
            return uri.Scheme == Uri.UriSchemeHttp && uri.IsLoopback;
        }

        public static bool IsConfigured =>
            !string.IsNullOrWhiteSpace(AppConfig.Instance?.CaseLibraryEndpoint);

        /// <summary>
        /// 连接案例库服务并返回独立的 MCP 管理器；调用方负责 Dispose。
        /// 连接失败时抛出 InvalidOperationException，消息可直接展示给用户。
        /// </summary>
        public static async Task<McpToolManager> ConnectAsync(CancellationToken cancellationToken)
        {
            AppConfig config = AppConfig.Instance;
            string endpoint = config?.CaseLibraryEndpoint?.Trim();
            if (string.IsNullOrEmpty(endpoint))
                throw new InvalidOperationException(Loc.T("I18n_CaseLibraryNotConfigured"));
            if (!IsValidEndpoint(endpoint))
                throw new InvalidOperationException(Loc.T("I18n_CaseLibraryEndpointInvalid"));

            var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string apiKey = config.CaseLibraryApiKey?.Trim();
            if (!string.IsNullOrEmpty(apiKey))
                headers["Authorization"] = "Bearer " + apiKey;

            var manager = new McpToolManager();
            try
            {
                await manager.AddServerAsync(
                    ServerName,
                    McpTransportKind.StreamableHttp,
                    endpoint,
                    httpHeaders: headers,
                    enabledToolNames: AllowedToolNames,
                    cancellationToken: cancellationToken,
                    trusted: true);

                if (manager.ServerCount == 0)
                    throw new InvalidOperationException(Loc.F("I18n_CaseLibraryConnectFailed", manager.LastError ?? ""));

                bool hasSearch = manager.GetEnabledTools().Any(t =>
                    t.ServerName == ServerName && t.Name == SearchToolName);
                if (!hasSearch)
                    throw new InvalidOperationException(Loc.F("I18n_CaseLibraryMissingTool", SearchToolName));

                return manager;
            }
            catch
            {
                manager.Dispose();
                throw;
            }
        }

        /// <summary>案例库启用时追加到 Agent 系统提示词末尾的使用规则。</summary>
        public static string BuildPromptRules(bool useEnglish)
        {
            return useEnglish
                ? @"

## Case Library (enabled for this task)
- `search_cases` searches the organization's library of administrative penalty and court cases; `get_case` returns one case by its case number.
- When reviewing a contract or advertising content, call `search_cases` for each distinct risk type you identify (e.g. truthfulness of content, review and verification duties, endorsements, allocation of liability, breach and recourse) before giving your opinion on it.
- Cite only case numbers and facts that appear in the tool results. Never invent cases. Put the case number in the comment or report.
- If a search returns nothing relevant, say ""no matching case in the library"". Never treat a missing case as evidence that there is no risk.
- Cases are risk references, not legal conclusions; penalty standards differ by region and period."
                : @"

## 案例库（本次任务已启用）
- `search_cases` 用于检索本单位的行政处罚与裁判案例库；`get_case` 按案号查看单个案例详情。
- 审查合同或广告内容时，每识别出一类风险（如内容真实性、审核查验义务、代言、责任分担、违约与追偿），先调用 `search_cases` 检索相关案例，再给出该项意见。
- 只能引用工具结果中出现的案号和事实，不得编造案例；在批注或报告中写明案号。
- 检索无相关结果时写明“案例库未命中”，不得据此认定不存在风险。
- 案例只作风险参考，不构成法律结论；不同地区、不同时期的处罚尺度可能不同。";
        }
    }

    /// <summary>
    /// 案例库检索工具。没有无参构造函数，因此不会被 ToolRegistry.DiscoverAndRegister 反射注册，
    /// 只能在用户勾选案例库后由 DocumentAgent.SetCaseLibrary 显式注册。
    /// </summary>
    public sealed class CaseLibraryTool : IAgentTool, ICancellableAgentTool
    {
        private readonly McpToolDefinition _def;
        private readonly McpToolManager _manager;

        public CaseLibraryTool(McpToolDefinition def, McpToolManager manager)
        {
            _def = def ?? throw new ArgumentNullException(nameof(def));
            _manager = manager ?? throw new ArgumentNullException(nameof(manager));
        }

        public string Name => _def.Name;
        public string Description => AgentToolText.F("I18n_AgentToolDescCaseLibrary", _def.Name, _def.Description ?? "");
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => true;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => _def.InputSchema ?? new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject(),
            ["required"] = new JArray()
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            return Task.Run(() => _manager.CallToolAsync(CaseLibrary.ServerName, _def.Name, parameters)).Result;
        }

        public Task<string> ExecuteAsync(JObject parameters, Word.Document doc, Word.Application app, CancellationToken cancellationToken)
        {
            return _manager.CallToolAsync(CaseLibrary.ServerName, _def.Name, parameters, cancellationToken);
        }
    }
}
