// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeCli.Core.Plugins;

/// <summary>
/// <see cref="IDocumentHandler"/> implementation that delegates every call to a
/// running format-handler plugin via <see cref="FormatHandlerSession"/>. Per
/// docs/plugin-protocol.md §2.3, this is what wraps the plugin so existing
/// get/view/query pipelines work transparently on foreign formats.
///
/// v0 scope: read-path methods (ViewAs*, Get, Query) and Validate are
/// proxied. Mutation methods throw a clear "not yet proxied" error. Future
/// versions will widen the surface as plugin authors need it.
/// </summary>
internal sealed class FormatHandlerProxy : IDocumentHandler
{
    private readonly FormatHandlerSession _session;

    public FormatHandlerProxy(FormatHandlerSession session) { _session = session; }

    // ----- Semantic layer (text views) -----------------------------------

    public string ViewAsText(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
        => SendViewString("text", startLine, endLine, maxLines, cols);

    public string ViewAsAnnotated(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
        => SendViewString("annotated", startLine, endLine, maxLines, cols);

    public string ViewAsOutline()
        => SendViewString("outline");

    public string ViewAsStats()
        => SendViewString("stats");

    public JsonNode ViewAsStatsJson() => SendViewJson("stats");
    public JsonNode ViewAsOutlineJson() => SendViewJson("outline");
    public JsonNode ViewAsTextJson(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
        => SendViewJson("text", startLine, endLine, maxLines, cols);

    public List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null)
    {
        var args = new JsonObject { ["mode"] = "issues" };
        if (issueType != null) args["type"] = issueType;
        if (limit.HasValue) args["limit"] = limit.Value;
        var result = _session.Send("command", "view", args);
        if (result is null) return new List<DocumentIssue>();
        return JsonSerializer.Deserialize(result.ToJsonString(), PluginJsonContext.Default.ListDocumentIssue) ?? new List<DocumentIssue>();
    }

    // ----- Query layer --------------------------------------------------

    public DocumentNode Get(string path, int depth = 1)
    {
        var result = _session.Send("command", "get", new JsonObject
        {
            ["path"] = path,
            ["depth"] = depth,
        });
        if (result is null)
            return new DocumentNode { Path = path, Type = "error", Text = "Plugin returned null result." };
        return JsonSerializer.Deserialize(result.ToJsonString(), PluginJsonContext.Default.DocumentNode)
            ?? new DocumentNode { Path = path, Type = "error", Text = "Plugin result did not deserialize as DocumentNode." };
    }

    public List<DocumentNode> Query(string selector)
    {
        var result = _session.Send("command", "query", new JsonObject { ["selector"] = selector });
        if (result is null) return new List<DocumentNode>();
        return JsonSerializer.Deserialize(result.ToJsonString(), PluginJsonContext.Default.ListDocumentNode) ?? new List<DocumentNode>();
    }

    public List<ValidationError> Validate()
    {
        var result = _session.Send("command", "validate", new JsonObject());
        if (result is null) return new List<ValidationError>();
        return JsonSerializer.Deserialize(result.ToJsonString(), PluginJsonContext.Default.ListValidationError) ?? new List<ValidationError>();
    }

    // ----- Mutation layer: v0 stubbed out -------------------------------

    public List<string> Set(string path, Dictionary<string, string> properties) => throw NotProxied("set");
    public string Add(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties) => throw NotProxied("add");
    public string? Remove(string path) => throw NotProxied("remove");
    public string Move(string sourcePath, string? targetParentPath, InsertPosition? position) => throw NotProxied("move");
    public string CopyFrom(string sourcePath, string targetParentPath, InsertPosition? position) => throw NotProxied("copy");
    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null) => throw NotProxied("raw");
    public void RawSet(string partPath, string xpath, string action, string? xml) => throw NotProxied("raw-set");
    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null) => throw NotProxied("add-part");

    public bool TryExtractBinary(string path, string destPath, out string? contentType, out long byteCount)
    {
        contentType = null;
        byteCount = 0;
        return false;
    }

    public void Dispose() => _session.Dispose();

    // ----- Helpers ------------------------------------------------------

    private string SendViewString(string mode, int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var args = BuildViewArgs(mode, startLine, endLine, maxLines, cols);
        var result = _session.Send("command", "view", args);
        return result?.GetValue<string>() ?? "";
    }

    private JsonNode SendViewJson(string mode, int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var args = BuildViewArgs(mode, startLine, endLine, maxLines, cols);
        args["format"] = "json";
        var result = _session.Send("command", "view", args);
        return result ?? new JsonObject();
    }

    private static JsonObject BuildViewArgs(string mode, int? startLine, int? endLine, int? maxLines, HashSet<string>? cols)
    {
        var args = new JsonObject { ["mode"] = mode };
        if (startLine.HasValue) args["start"] = startLine.Value;
        if (endLine.HasValue) args["end"] = endLine.Value;
        if (maxLines.HasValue) args["max-lines"] = maxLines.Value;
        if (cols != null && cols.Count > 0) args["cols"] = string.Join(",", cols);
        return args;
    }

    private static CliException NotProxied(string op) =>
        new CliException($"Format-handler v0 does not yet proxy `{op}`. Only view/get/query are wired through.")
        { Code = "format_handler_not_supported" };
}
