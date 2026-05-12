// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Diagnostics;
using OfficeCli.Core;
using OfficeCli.Core.Plugins;

namespace OfficeCli;

static partial class CommandBuilder
{
    /// <summary>
    /// `officecli export <file> --to <ext> [--out <path>]` — convert a native
    /// document to a foreign format via an installed exporter plugin.
    /// Subprocess-only protocol: no IPC, no commands, the plugin reads the
    /// source and writes the target then exits.
    /// </summary>
    private static Command BuildExportCommand(Option<bool> jsonOption)
    {
        var fileArg = new Argument<FileInfo>("file") { Description = "Source document (.docx/.xlsx/.pptx)" };
        var toOpt = new Option<string>("--to") { Description = "Target format extension (e.g. 'pdf', 'epub')", Required = true };
        var outOpt = new Option<string?>("--out", "-o") { Description = "Target file path (default: source path with new extension)" };

        var cmd = new Command("export", "Convert a document to another format using an exporter plugin");
        cmd.Add(fileArg);
        cmd.Add(toOpt);
        cmd.Add(outOpt);
        cmd.Add(jsonOption);

        cmd.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(fileArg)!;
            var toRaw = result.GetValue(toOpt) ?? throw new CliException("--to is required") { Code = "missing_argument" };
            var outPath = result.GetValue(outOpt);

            var sourceExt = Path.GetExtension(file.FullName).ToLowerInvariant();
            var targetExt = NormalizeExt(toRaw);

            if (string.IsNullOrEmpty(targetExt) || targetExt == ".")
                throw new CliException($"Invalid --to value: '{toRaw}'") { Code = "invalid_argument" };

            // Resolve the exporter plugin. We look it up by *target* extension —
            // plugin manifests for exporters declare what format they produce
            // (e.g. `extensions: [".pdf"]`), and the source format is matched
            // via the manifest's `supports` list (`from:docx`).
            var plugin = ResolveExporter(sourceExt, targetExt);
            if (plugin is null)
                throw new CliException(
                    $"No exporter plugin found for {sourceExt} → {targetExt}.")
                {
                    Code = "exporter_not_found",
                    Suggestion = "Install an exporter plugin: `officecli plugins list` to see what's available, or see docs/plugin-protocol.md.",
                };

            outPath ??= Path.ChangeExtension(file.FullName, targetExt.TrimStart('.'));

            // If a resident holds the file open, it has an exclusive lock and
            // any exporter (Word interop, soffice, ...) will fail with "file
            // in use". Ask the resident to flush+close before we snapshot.
            // The resident can be reopened by the user afterwards if needed.
            bool residentClosed = false;
            if (ResidentClient.TryConnect(file.FullName, out _))
            {
                if (ResidentClient.SendCloseWithResponse(file.FullName, out _))
                    residentClosed = true;
            }

            // Snapshot the source to a tmp path so the plugin can open it
            // exclusively. Safe because exporters are read-only with respect
            // to the source.
            var pluginSource = SnapshotForExport(file.FullName);
            try
            {
                var (exitCode, stderr) = RunExporter(plugin.ExecutablePath, pluginSource, outPath);
                if (exitCode != 0)
                    throw new CliException(
                        $"Exporter plugin '{plugin.Manifest.Name}' failed (exit {exitCode}): {Truncate(stderr, 500)}")
                    {
                        Code = exitCode switch
                        {
                            2 => "corrupt_input",
                            3 => "unsupported_feature",
                            4 => "license_expired",
                            5 => "protocol_mismatch",
                            _ => "plugin_failed",
                        },
                    };

                if (!File.Exists(outPath))
                    throw new CliException(
                        $"Exporter plugin '{plugin.Manifest.Name}' reported success but no output file was written at {outPath}.")
                    { Code = "plugin_contract_violation" };

                var residentNote = residentClosed
                    ? " (resident closed to release lock; reopen with `officecli open` if needed)"
                    : "";
                var msg = $"Exported: {file.FullName} → {outPath} (plugin: {plugin.Manifest.Name}){residentNote}";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
                else Console.WriteLine(msg);
                return 0;
            }
            finally
            {
                try { File.Delete(pluginSource); } catch { }
            }
        }, json); });

        return cmd;
    }

    /// <summary>
    /// Find an exporter for (source, target). Indexes by target extension (the
    /// plugin's declared extensions field); filters by source via the manifest's
    /// supports list. A plugin missing supports is assumed to accept all native
    /// sources — conservative default for older manifests.
    /// </summary>
    private static ResolvedPlugin? ResolveExporter(string sourceExt, string targetExt)
    {
        var p = PluginRegistry.FindFor(PluginKind.Exporter, targetExt);
        if (p is null) return null;

        if (p.Manifest.Supports is null || p.Manifest.Supports.Count == 0)
            return p; // legacy/permissive

        var sourceBare = sourceExt.TrimStart('.');
        if (p.Manifest.Supports.Any(s =>
                string.Equals(s, $"from:{sourceBare}", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(s, sourceExt, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(s, sourceBare, StringComparison.OrdinalIgnoreCase)))
            return p;

        return null;
    }

    /// <summary>
    /// Copy the source to a fresh tmp file so the exporter plugin can open it
    /// exclusively, even when a resident/Word/another process already holds the
    /// original. Uses FileShare.ReadWrite on the read side so the snapshot
    /// succeeds even against an exclusive write lock; the produced tmp file is
    /// the plugin's input. Caller is responsible for deleting it.
    /// </summary>
    private static string SnapshotForExport(string sourcePath)
    {
        var ext = Path.GetExtension(sourcePath);
        var tmp = Path.Combine(Path.GetTempPath(),
            $"officecli-export-{Guid.NewGuid():N}{ext}");
        using var src = new FileStream(sourcePath, FileMode.Open, FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete);
        using var dst = new FileStream(tmp, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        src.CopyTo(dst);
        return tmp;
    }

    private static (int exitCode, string stderr) RunExporter(string exe, string source, string target)
    {
        using var p = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = exe,
                ArgumentList = { "export", source, "--out", target },
                UseShellExecute = false,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                CreateNoWindow = true,
            }
        };
        p.Start();
        // Drain stdout so the plugin doesn't block on a full pipe; keep stderr
        // for error surfacing.
        _ = p.StandardOutput.ReadToEndAsync();
        var stderrTask = p.StandardError.ReadToEndAsync();
        p.WaitForExit();
        return (p.ExitCode, stderrTask.Result);
    }

    private static string NormalizeExt(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";
        if (!raw.StartsWith('.')) raw = "." + raw;
        return raw.ToLowerInvariant();
    }
}
