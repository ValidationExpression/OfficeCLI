// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Help;

namespace OfficeCli;

static partial class CommandBuilder
{
    /// <summary>
    /// `officecli help [format] [element] [--json]` — schema-driven help.
    /// Single entry point replacing the old per-format docx/xlsx/pptx drill-down.
    /// </summary>
    public static Command BuildHelpCommand(Option<bool> jsonOption)
    {
        var formatArg = new Argument<string?>("format")
        {
            Description = "Document format: docx/xlsx/pptx (aliases: word, excel, ppt, powerpoint). Omit to list formats.",
            Arity = ArgumentArity.ZeroOrOne,
        };
        var elementArg = new Argument<string?>("element")
        {
            Description = "Element name (e.g. paragraph, cell, shape). Omit to list elements for the format.",
            Arity = ArgumentArity.ZeroOrOne,
        };

        var command = new Command("help", "Show schema-driven capability reference for officecli.");
        command.Add(formatArg);
        command.Add(elementArg);
        command.Add(jsonOption);

        command.SetAction(result =>
        {
            var json = result.GetValue(jsonOption);
            var format = result.GetValue(formatArg);
            var element = result.GetValue(elementArg);
            return SafeRun(() => RunHelp(format, element, json), json);
        });

        return command;
    }

    private static int RunHelp(string? format, string? element, bool json)
    {
        // Case 1: no args — list formats and usage banner.
        if (string.IsNullOrEmpty(format))
        {
            Console.WriteLine("officecli help — capability reference (schema-driven)");
            Console.WriteLine();
            Console.WriteLine("Formats:");
            foreach (var f in SchemaHelpLoader.ListFormats())
                Console.WriteLine($"  {f}");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  officecli help <format>                   List elements");
            Console.WriteLine("  officecli help <format> <element>         Human-readable detail");
            Console.WriteLine("  officecli help <format> <element> --json  Raw schema JSON");
            Console.WriteLine();
            Console.WriteLine("Aliases: word→docx, excel→xlsx, ppt/powerpoint→pptx");
            return 0;
        }

        // Case 2: format only — list elements.
        if (string.IsNullOrEmpty(element))
        {
            var canonical = SchemaHelpLoader.NormalizeFormat(format);
            var elements = SchemaHelpLoader.ListElements(canonical);
            Console.WriteLine($"Elements for {canonical}:");
            foreach (var el in elements)
                Console.WriteLine($"  {el}");
            Console.WriteLine();
            Console.WriteLine($"Run 'officecli help {canonical} <element>' for detail.");
            return 0;
        }

        // Case 3: format + element — render schema.
        using var doc = SchemaHelpLoader.LoadSchema(format, element);
        Console.WriteLine(json
            ? SchemaHelpRenderer.RenderJson(doc)
            : SchemaHelpRenderer.RenderHuman(doc));
        return 0;
    }
}
