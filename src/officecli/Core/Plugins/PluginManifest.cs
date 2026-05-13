// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json.Serialization;
using OfficeCli.Core;

namespace OfficeCli.Core.Plugins;

/// <summary>
/// The three plugin responsibilities defined in docs/plugin-protocol.md.
/// String values are the wire form used in plugin manifests.
/// </summary>
public enum PluginKind
{
    /// <summary>Foreign format → officecli commands (e.g. .doc → .docx via add/set).</summary>
    DumpReader,

    /// <summary>Native format → foreign output file (e.g. .docx → .pdf).</summary>
    Exporter,

    /// <summary>Plugin owns a foreign format end-to-end (e.g. .hwpx editing).</summary>
    FormatHandler,
}

public static class PluginKindExtensions
{
    public static string ToWireString(this PluginKind kind) => kind switch
    {
        PluginKind.DumpReader    => "dump-reader",
        PluginKind.Exporter      => "exporter",
        PluginKind.FormatHandler => "format-handler",
        _ => throw new ArgumentOutOfRangeException(nameof(kind)),
    };

    public static bool TryParseWire(string s, out PluginKind kind)
    {
        switch (s)
        {
            case "dump-reader":    kind = PluginKind.DumpReader;    return true;
            case "exporter":       kind = PluginKind.Exporter;      return true;
            case "format-handler": kind = PluginKind.FormatHandler; return true;
            default:               kind = default;                  return false;
        }
    }
}

/// <summary>
/// Manifest emitted by a plugin in response to `<plugin> --info`. Mirrors
/// the schema defined in docs/plugin-protocol.md §4.
/// </summary>
public sealed class PluginManifest
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("version")]
    public string Version { get; set; } = "";

    [JsonPropertyName("protocol")]
    public int Protocol { get; set; }

    /// <summary>Wire-form kind strings (`"dump-reader"`, etc.). Parsed via <see cref="PluginKindExtensions.TryParseWire"/>.</summary>
    [JsonPropertyName("kinds")]
    public List<string> Kinds { get; set; } = new();

    /// <summary>File extensions including the leading dot (e.g. `".doc"`).</summary>
    [JsonPropertyName("extensions")]
    public List<string> Extensions { get; set; } = new();

    /// <summary>
    /// Native format the plugin produces (dump-reader: the format the emitted
    /// batch is replayed into; exporter: the source-side native format).
    /// One of <c>"docx"</c>, <c>"xlsx"</c>, <c>"pptx"</c>. Optional — defaults
    /// to <c>"docx"</c> for protocol-v1 plugins that pre-date this field.
    /// Used by main to pick the sibling-cache extension, the linting schema,
    /// and the handler that opens replay output.
    /// </summary>
    [JsonPropertyName("target")]
    public string? Target { get; set; }

    [JsonPropertyName("description")]
    public string? Description { get; set; }

    [JsonPropertyName("tier")]
    public string? Tier { get; set; }

    [JsonPropertyName("supports")]
    public List<string>? Supports { get; set; }

    [JsonPropertyName("limits")]
    public Dictionary<string, object>? Limits { get; set; }

    [JsonPropertyName("homepage")]
    public string? Homepage { get; set; }

    [JsonPropertyName("license")]
    public string? License { get; set; }

    [JsonPropertyName("vocabulary")]
    public PluginVocabulary? Vocabulary { get; set; }
}

public static class PluginManifestExtensions
{
    /// <summary>
    /// Canonical target format name ("docx"/"xlsx"/"pptx"). Defaults to
    /// "docx" for plugins that omit the field (protocol-v1 back-compat).
    /// Throws if the manifest declares an unsupported target.
    /// </summary>
    public static string ResolveTargetFormat(this PluginManifest m)
    {
        var t = (m.Target ?? "docx").ToLowerInvariant();
        return t switch
        {
            "docx" or "xlsx" or "pptx" => t,
            _ => throw new InvalidOperationException(
                $"Plugin '{m.Name}' declares unsupported target '{m.Target}'. Expected one of: docx, xlsx, pptx."),
        };
    }

    /// <summary>
    /// File extension (with leading dot) for the plugin's target format.
    /// </summary>
    public static string ResolveTargetExtension(this PluginManifest m) =>
        "." + m.ResolveTargetFormat();
}

/// <summary>
/// Format-handler plugins declare the document model they expose via this
/// vocabulary. Used by main for autocomplete, command validation, and help.
/// Main does not interpret the semantics — it forwards commands using these names.
/// </summary>
public sealed class PluginVocabulary
{
    [JsonPropertyName("addable_types")]
    public List<string> AddableTypes { get; set; } = new();

    /// <summary>Map from type name (e.g. `"page"`) to the property names that type accepts.</summary>
    [JsonPropertyName("settable_props")]
    public Dictionary<string, List<string>> SettableProps { get; set; } = new();

    [JsonPropertyName("path_segments")]
    public List<string> PathSegments { get; set; } = new();
}

[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.SnakeCaseLower,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull)]
[JsonSerializable(typeof(PluginManifest))]
[JsonSerializable(typeof(PluginVocabulary))]
[JsonSerializable(typeof(List<string>))]
[JsonSerializable(typeof(DocumentNode))]
[JsonSerializable(typeof(List<DocumentNode>))]
[JsonSerializable(typeof(DocumentIssue))]
[JsonSerializable(typeof(List<DocumentIssue>))]
[JsonSerializable(typeof(ValidationError))]
[JsonSerializable(typeof(List<ValidationError>))]
internal partial class PluginJsonContext : JsonSerializerContext;
