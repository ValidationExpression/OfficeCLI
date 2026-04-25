// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

// Per-element-type Set helpers for Excel paths. Mechanically extracted from
// the original god-method Set(); each helper owns one path-pattern's full
// handling. No behavior change.
public partial class ExcelHandler
{
    private List<string> SetSparklineByPath(Match m, Dictionary<string, string> properties)
    {
        var spkSheet = m.Groups[1].Value;
        var spkIdx = int.Parse(m.Groups[2].Value);
        var spkWorksheet = FindWorksheet(spkSheet) ?? throw SheetNotFoundException(spkSheet);
        var spkGroup = GetSparklineGroup(spkWorksheet, spkIdx)
            ?? throw new ArgumentException($"Sparkline[{spkIdx}] not found in sheet '{spkSheet}'");

        var unsup = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "type":
                    spkGroup.Type = value.ToLowerInvariant() switch
                    {
                        "column" => X14.SparklineTypeValues.Column,
                        "stacked" => X14.SparklineTypeValues.Stacked,
                        _ => null
                    };
                    break;
                case "color":
                    spkGroup.SeriesColor = new X14.SeriesColor { Rgb = ParseHelpers.NormalizeArgbColor(value) };
                    break;
                case "negativecolor":
                    spkGroup.NegativeColor = new X14.NegativeColor { Rgb = ParseHelpers.NormalizeArgbColor(value) };
                    break;
                case "markers":
                    spkGroup.Markers = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                    break;
                case "highpoint":
                    spkGroup.High = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                    break;
                case "lowpoint":
                    spkGroup.Low = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                    break;
                case "firstpoint":
                    spkGroup.First = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                    break;
                case "lastpoint":
                    spkGroup.Last = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                    break;
                case "negative":
                    spkGroup.Negative = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                    break;
                case "lineweight":
                    if (double.TryParse(value, out var lw)) spkGroup.LineWeight = lw;
                    break;
                case "datarange" or "range":
                {
                    var newRangeRef = value.Contains('!') ? value : $"{spkSheet}!{value}";
                    foreach (var spk in spkGroup.Descendants<X14.Sparkline>())
                    {
                        var f = spk.GetFirstChild<DocumentFormat.OpenXml.Office.Excel.Formula>();
                        if (f != null) f.Text = newRangeRef;
                        else spk.InsertAt(new DocumentFormat.OpenXml.Office.Excel.Formula(newRangeRef), 0);
                    }
                    break;
                }
                case "location" or "cell":
                {
                    foreach (var spk in spkGroup.Descendants<X14.Sparkline>())
                    {
                        var r = spk.GetFirstChild<DocumentFormat.OpenXml.Office.Excel.ReferenceSequence>();
                        if (r != null) r.Text = value;
                        else spk.AppendChild(new DocumentFormat.OpenXml.Office.Excel.ReferenceSequence(value));
                    }
                    break;
                }
                default:
                    unsup.Add(key);
                    break;
            }
        }
        SaveWorksheet(spkWorksheet);
        return unsup;
    }

    private List<string> SetNamedRangeByPath(Match m, Dictionary<string, string> properties)
    {
        var selector = m.Groups[1].Value;
        var workbook = GetWorkbook();
        var definedNames = workbook.GetFirstChild<DefinedNames>()
            ?? throw new ArgumentException("No named ranges found in workbook");

        var allDefs = definedNames.Elements<DefinedName>().ToList();
        DefinedName? dn;

        if (int.TryParse(selector, out var dnIndex))
        {
            if (dnIndex < 1 || dnIndex > allDefs.Count)
                throw new ArgumentException($"Named range index {dnIndex} out of range (1-{allDefs.Count})");
            dn = allDefs[dnIndex - 1];
        }
        else
        {
            dn = allDefs.FirstOrDefault(d =>
                d.Name?.Value?.Equals(selector, StringComparison.OrdinalIgnoreCase) == true)
                ?? throw new ArgumentException($"Named range '{selector}' not found");
        }

        var nrUnsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "ref": dn.Text = value; break;
                case "name": dn.Name = value; break;
                case "comment": dn.Comment = value; break;
                case "scope":
                    if (string.IsNullOrEmpty(value) || value.Equals("workbook", StringComparison.OrdinalIgnoreCase))
                    {
                        dn.LocalSheetId = null;
                    }
                    else
                    {
                        var nrSheets = workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToList();
                        var nrSheetIdx = nrSheets?.FindIndex(s =>
                            s.Name?.Value?.Equals(value, StringComparison.OrdinalIgnoreCase) == true) ?? -1;
                        if (nrSheetIdx >= 0)
                            dn.LocalSheetId = (uint)nrSheetIdx;
                        else
                            throw new ArgumentException($"Sheet '{value}' not found for scope");
                    }
                    break;
                default: nrUnsupported.Add(key); break;
            }
        }

        workbook.Save();
        return nrUnsupported;
    }

    private List<string> SetValidationByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var dvIdx = int.Parse(m.Groups[1].Value);
        var dvs = GetSheet(worksheet).GetFirstChild<DataValidations>()
            ?? throw new ArgumentException("No data validations found in sheet");

        var dvList = dvs.Elements<DataValidation>().ToList();
        if (dvIdx < 1 || dvIdx > dvList.Count)
            throw new ArgumentException($"Validation index {dvIdx} out of range (1-{dvList.Count})");

        var dv = dvList[dvIdx - 1];
        var dvUnsupported = new List<string>();

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                // CONSISTENCY(canonical-key): schema canonical key is 'ref';
                // 'sqref' retained as legacy alias.
                case "sqref" or "ref":
                    dv.SequenceOfReferences = new ListValue<StringValue>(
                        value.Split(' ').Select(s => new StringValue(s)));
                    break;
                case "type":
                    dv.Type = value.ToLowerInvariant() switch
                    {
                        "list" => DataValidationValues.List,
                        "whole" => DataValidationValues.Whole,
                        "decimal" => DataValidationValues.Decimal,
                        "date" => DataValidationValues.Date,
                        "time" => DataValidationValues.Time,
                        "textlength" => DataValidationValues.TextLength,
                        "custom" => DataValidationValues.Custom,
                        _ => throw new ArgumentException($"Unknown validation type: '{value}'. Valid types: list, whole, decimal, date, time, textLength, custom.")
                    };
                    break;
                case "formula1":
                    if (dv.Type?.Value == DataValidationValues.List && !value.StartsWith("\""))
                        dv.Formula1 = new Formula1($"\"{value}\"");
                    else
                        dv.Formula1 = new Formula1(value);
                    break;
                case "formula2":
                    dv.Formula2 = new Formula2(value);
                    break;
                case "operator":
                    dv.Operator = value.ToLowerInvariant() switch
                    {
                        "between" => DataValidationOperatorValues.Between,
                        "notbetween" => DataValidationOperatorValues.NotBetween,
                        "equal" => DataValidationOperatorValues.Equal,
                        "notequal" => DataValidationOperatorValues.NotEqual,
                        "lessthan" => DataValidationOperatorValues.LessThan,
                        "lessthanorequal" => DataValidationOperatorValues.LessThanOrEqual,
                        "greaterthan" => DataValidationOperatorValues.GreaterThan,
                        "greaterthanorequal" => DataValidationOperatorValues.GreaterThanOrEqual,
                        _ => throw new ArgumentException($"Unknown operator: {value}")
                    };
                    break;
                case "allowblank": dv.AllowBlank = IsTruthy(value); break;
                case "showerror": dv.ShowErrorMessage = IsTruthy(value); break;
                case "errortitle": dv.ErrorTitle = value; break;
                case "error": dv.Error = value; break;
                case "showinput": dv.ShowInputMessage = IsTruthy(value); break;
                case "prompttitle": dv.PromptTitle = value; break;
                case "prompt": dv.Prompt = value; break;
                default: dvUnsupported.Add(key); break;
            }
        }

        SaveWorksheet(worksheet);
        return dvUnsupported;
    }

    // Replace backing embedded part + refresh ProgID. Cleans up the old payload
    // part (CLAUDE.md Known API Quirks rule: always delete the old part on src
    // replacement).
    private List<string> SetOleByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var oleIdxSet = int.Parse(m.Groups[1].Value);
        var oleWs = GetSheet(worksheet);
        var oleElements = oleWs.Descendants<OleObject>().ToList();
        if (oleIdxSet < 1 || oleIdxSet > oleElements.Count)
            throw new ArgumentException($"OLE object index {oleIdxSet} out of range (1..{oleElements.Count})");
        var oleObjSet = oleElements[oleIdxSet - 1];
        var oleUnsupportedSet = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "path" or "src":
                {
                    if (oleObjSet.Id?.Value is string oldRel && !string.IsNullOrEmpty(oldRel))
                    {
                        try { worksheet.DeletePart(oldRel); } catch { }
                    }
                    var (newRel, _) = OfficeCli.Core.OleHelper.AddEmbeddedPart(worksheet, value, _filePath);
                    oleObjSet.Id = newRel;
                    if (!properties.ContainsKey("progId") && !properties.ContainsKey("progid"))
                    {
                        var autoProgId = OfficeCli.Core.OleHelper.DetectProgId(value);
                        OfficeCli.Core.OleHelper.ValidateProgId(autoProgId);
                        oleObjSet.ProgId = autoProgId;
                    }
                    break;
                }
                case "progid":
                    OfficeCli.Core.OleHelper.ValidateProgId(value);
                    oleObjSet.ProgId = value;
                    break;
                case "display":
                    // CONSISTENCY(excel-ole-display): Excel Add rejects 'display'
                    // with ArgumentException; Set must do the same instead of
                    // falling into the default unsupported branch.
                    throw new ArgumentException(
                        "'display' property is not supported for Excel OLE "
                        + "(Excel always shows objects as icon). Remove --prop display.");
                case "width":
                case "height":
                {
                    // CONSISTENCY(ole-width-units): accept either bare integer cell-span or unit-qualified size.
                    long emuTotal;
                    try { emuTotal = ParseAnchorDimensionEmu(value, key.ToLowerInvariant()); }
                    catch { oleUnsupportedSet.Add(key); break; }
                    if (emuTotal < 0) { oleUnsupportedSet.Add(key); break; }
                    var objectPrSet = oleObjSet.GetFirstChild<EmbeddedObjectProperties>();
                    var objAnchorSet = objectPrSet?.GetFirstChild<ObjectAnchor>();
                    var fromMSet = objAnchorSet?.GetFirstChild<FromMarker>();
                    var toMSet = objAnchorSet?.GetFirstChild<ToMarker>();
                    if (fromMSet == null || toMSet == null) { oleUnsupportedSet.Add(key); break; }
                    if (key.Equals("width", StringComparison.OrdinalIgnoreCase))
                    {
                        int.TryParse(fromMSet.GetFirstChild<XDR.ColumnId>()?.Text ?? "0", out var fromCol);
                        long.TryParse(fromMSet.GetFirstChild<XDR.ColumnOffset>()?.Text ?? "0", out var fromColOff);
                        long wholeCols = emuTotal / EmuPerColApprox;
                        long remCols = emuTotal % EmuPerColApprox;
                        var toColChild = toMSet.GetFirstChild<XDR.ColumnId>();
                        if (toColChild != null) toColChild.Text = (fromCol + (int)wholeCols).ToString();
                        var toColOffChild = toMSet.GetFirstChild<XDR.ColumnOffset>();
                        if (toColOffChild != null) toColOffChild.Text = (fromColOff + remCols).ToString();
                        else toMSet.InsertAfter(new XDR.ColumnOffset((fromColOff + remCols).ToString()), toColChild);
                    }
                    else
                    {
                        int.TryParse(fromMSet.GetFirstChild<XDR.RowId>()?.Text ?? "0", out var fromRow);
                        long.TryParse(fromMSet.GetFirstChild<XDR.RowOffset>()?.Text ?? "0", out var fromRowOff);
                        long wholeRows = emuTotal / EmuPerRowApprox;
                        long remRows = emuTotal % EmuPerRowApprox;
                        var toRowChild = toMSet.GetFirstChild<XDR.RowId>();
                        if (toRowChild != null) toRowChild.Text = (fromRow + (int)wholeRows).ToString();
                        var toRowOffChild = toMSet.GetFirstChild<XDR.RowOffset>();
                        if (toRowOffChild != null) toRowOffChild.Text = (fromRowOff + remRows).ToString();
                        else toMSet.InsertAfter(new XDR.RowOffset((fromRowOff + remRows).ToString()), toRowChild);
                    }
                    break;
                }
                case "anchor":
                {
                    // CONSISTENCY(ole-width-units): mirror Add-side warn — width/height
                    // dropped silently when anchor= present.
                    if (properties.ContainsKey("width") || properties.ContainsKey("height"))
                        Console.Error.WriteLine(
                            "Warning: 'width'/'height' are ignored when 'anchor' is provided (anchor defines the full rectangle).");
                    var anchorM = Regex.Match(value ?? "", @"^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$", RegexOptions.IgnoreCase);
                    if (!anchorM.Success) { oleUnsupportedSet.Add(key); break; }
                    var objectPrAnc = oleObjSet.GetFirstChild<EmbeddedObjectProperties>();
                    var objAnchorAnc = objectPrAnc?.GetFirstChild<ObjectAnchor>();
                    var fromMAnc = objAnchorAnc?.GetFirstChild<FromMarker>();
                    var toMAnc = objAnchorAnc?.GetFirstChild<ToMarker>();
                    if (fromMAnc == null || toMAnc == null) { oleUnsupportedSet.Add(key); break; }
                    int newFromCol = ColumnNameToIndex(anchorM.Groups[1].Value) - 1;
                    int newFromRow = int.Parse(anchorM.Groups[2].Value) - 1;
                    int newToCol, newToRow;
                    if (anchorM.Groups[3].Success)
                    {
                        newToCol = ColumnNameToIndex(anchorM.Groups[3].Value) - 1;
                        newToRow = int.Parse(anchorM.Groups[4].Value) - 1;
                    }
                    else
                    {
                        newToCol = newFromCol + 2;
                        newToRow = newFromRow + 3;
                    }
                    var fromColChild = fromMAnc.GetFirstChild<XDR.ColumnId>();
                    if (fromColChild != null) fromColChild.Text = newFromCol.ToString();
                    var fromRowChild = fromMAnc.GetFirstChild<XDR.RowId>();
                    if (fromRowChild != null) fromRowChild.Text = newFromRow.ToString();
                    var fromColOffChild = fromMAnc.GetFirstChild<XDR.ColumnOffset>();
                    if (fromColOffChild != null) fromColOffChild.Text = "0";
                    var fromRowOffChild = fromMAnc.GetFirstChild<XDR.RowOffset>();
                    if (fromRowOffChild != null) fromRowOffChild.Text = "0";
                    var toColChildAnc = toMAnc.GetFirstChild<XDR.ColumnId>();
                    if (toColChildAnc != null) toColChildAnc.Text = newToCol.ToString();
                    var toRowChildAnc = toMAnc.GetFirstChild<XDR.RowId>();
                    if (toRowChildAnc != null) toRowChildAnc.Text = newToRow.ToString();
                    var toColOffChildAnc = toMAnc.GetFirstChild<XDR.ColumnOffset>();
                    if (toColOffChildAnc != null) toColOffChildAnc.Text = "0";
                    var toRowOffChildAnc = toMAnc.GetFirstChild<XDR.RowOffset>();
                    if (toRowOffChildAnc != null) toRowOffChildAnc.Text = "0";
                    break;
                }
                default:
                    oleUnsupportedSet.Add(key);
                    break;
            }
        }
        SaveWorksheet(worksheet);
        return oleUnsupportedSet;
    }

    private List<string> SetPictureByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var picIdx = int.Parse(m.Groups[1].Value);
        var drawingsPart = worksheet.DrawingsPart
            ?? throw new ArgumentException("Sheet has no drawings/pictures");
        var wsDrawing = drawingsPart.WorksheetDrawing
            ?? throw new ArgumentException("Sheet has no drawings/pictures");

        var picAnchors = wsDrawing.Elements<XDR.TwoCellAnchor>()
            .Where(a => a.Descendants<XDR.Picture>().Any()).ToList();
        if (picIdx < 1 || picIdx > picAnchors.Count)
            throw new ArgumentException($"Picture index {picIdx} out of range (1..{picAnchors.Count})");

        var anchor = picAnchors[picIdx - 1];
        var picUnsupported = new List<string>();

        foreach (var (key, value) in properties)
        {
            var lk = key.ToLowerInvariant();
            if (TrySetAnchorPosition(anchor, lk, value)) continue;

            var spPr = anchor.Descendants<XDR.ShapeProperties>().FirstOrDefault();
            if (TrySetRotation(spPr, lk, value)) continue;
            if (TrySetShapeFlip(spPr, lk, value)) continue;
            if (TrySetShapeEffect(spPr, lk, value)) continue;

            switch (lk)
            {
                case "alt":
                    var nvProps = anchor.Descendants<XDR.NonVisualDrawingProperties>().FirstOrDefault();
                    if (nvProps != null) nvProps.Description = value;
                    break;
                default:
                    picUnsupported.Add(key);
                    break;
            }
        }

        drawingsPart.WorksheetDrawing.Save();
        return picUnsupported;
    }

    private List<string> SetShapeByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var shpIdx = int.Parse(m.Groups[1].Value);
        var drawingsPart = worksheet.DrawingsPart
            ?? throw new ArgumentException("Sheet has no drawings/shapes");
        var wsDrawing = drawingsPart.WorksheetDrawing
            ?? throw new ArgumentException("Sheet has no drawings/shapes");

        var shpAnchors = wsDrawing.Elements<XDR.TwoCellAnchor>()
            .Where(a => a.Descendants<XDR.Shape>().Any()).ToList();
        if (shpIdx < 1 || shpIdx > shpAnchors.Count)
            throw new ArgumentException($"Shape index {shpIdx} out of range (1..{shpAnchors.Count})");

        var anchor = shpAnchors[shpIdx - 1];
        var shape = anchor.Descendants<XDR.Shape>().First();
        var shpUnsupported = new List<string>();

        foreach (var (key, value) in properties)
        {
            var lk = key.ToLowerInvariant();
            if (TrySetAnchorPosition(anchor, lk, value)) continue;
            if (TrySetRotation(shape.ShapeProperties, lk, value)) continue;
            if (TrySetShapeFlip(shape.ShapeProperties, lk, value)) continue;
            if (TrySetShapeFontProp(shape, lk, value)) continue;

            // For effects on shapes: check if fill=none → text-level, otherwise shape-level
            if (lk is "shadow" or "glow" or "reflection" or "softedge")
            {
                var spPr = shape.ShapeProperties;
                if (spPr == null) continue;
                var isNoFill = spPr.GetFirstChild<Drawing.NoFill>() != null;
                var normalizedVal = value.Replace(':', '-');

                if (isNoFill && lk is "shadow" or "glow")
                {
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        if (lk == "shadow")
                            OfficeCli.Core.DrawingEffectsHelper.ApplyTextEffect<Drawing.OuterShadow>(run, normalizedVal, () =>
                                OfficeCli.Core.DrawingEffectsHelper.BuildOuterShadow(normalizedVal, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor));
                        else
                            OfficeCli.Core.DrawingEffectsHelper.ApplyTextEffect<Drawing.Glow>(run, normalizedVal, () =>
                                OfficeCli.Core.DrawingEffectsHelper.BuildGlow(normalizedVal, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor));
                    }
                }
                else
                {
                    TrySetShapeEffect(spPr, lk, value);
                }
                continue;
            }

            switch (lk)
            {
                case "name":
                {
                    var nvProps = shape.NonVisualShapeProperties?.GetFirstChild<XDR.NonVisualDrawingProperties>();
                    if (nvProps != null) nvProps.Name = value;
                    break;
                }
                case "text":
                {
                    var txBody = shape.TextBody;
                    if (txBody != null)
                    {
                        var firstPara = txBody.Elements<Drawing.Paragraph>().FirstOrDefault();
                        var pProps = firstPara?.ParagraphProperties?.CloneNode(true);
                        var rProps = firstPara?.Elements<Drawing.Run>().FirstOrDefault()?.RunProperties?.CloneNode(true);
                        txBody.RemoveAllChildren<Drawing.Paragraph>();
                        var lines = value.Replace("\\n", "\n").Split('\n');
                        foreach (var line in lines)
                        {
                            var para = new Drawing.Paragraph();
                            if (pProps != null) para.AppendChild(pProps.CloneNode(true));
                            var run = new Drawing.Run(new Drawing.Text(line));
                            if (rProps != null) run.RunProperties = (Drawing.RunProperties)rProps.CloneNode(true);
                            para.AppendChild(run);
                            txBody.AppendChild(para);
                        }
                    }
                    break;
                }
                case "font":
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rPr.RemoveAllChildren<Drawing.LatinFont>();
                        rPr.RemoveAllChildren<Drawing.EastAsianFont>();
                        rPr.AppendChild(new Drawing.LatinFont { Typeface = value });
                        rPr.AppendChild(new Drawing.EastAsianFont { Typeface = value });
                    }
                    break;
                case "size":
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rPr.FontSize = (int)Math.Round(ParseHelpers.ParseFontSize(value) * 100);
                    }
                    break;
                case "bold":
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rPr.Bold = IsTruthy(value);
                    }
                    break;
                case "italic":
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rPr.Italic = IsTruthy(value);
                    }
                    break;
                case "color":
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rPr.RemoveAllChildren<Drawing.SolidFill>();
                        var (cRgb, _) = ParseHelpers.SanitizeColorForOoxml(value);
                        OfficeCli.Core.DrawingEffectsHelper.InsertFillInRunProperties(rPr,
                            new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = cRgb }));
                    }
                    break;
                case "fill":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr != null)
                    {
                        spPr.RemoveAllChildren<Drawing.SolidFill>();
                        spPr.RemoveAllChildren<Drawing.NoFill>();
                        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                            spPr.AppendChild(new Drawing.NoFill());
                        else
                        {
                            var (fRgb, _) = ParseHelpers.SanitizeColorForOoxml(value);
                            spPr.AppendChild(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = fRgb }));
                        }
                    }
                    break;
                }
                case "align":
                    foreach (var para in shape.Descendants<Drawing.Paragraph>())
                    {
                        var pPr = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pPr.Alignment = value.ToLowerInvariant() switch
                        {
                            "center" or "c" or "ctr" => Drawing.TextAlignmentTypeValues.Center,
                            "right" or "r" => Drawing.TextAlignmentTypeValues.Right,
                            "justify" or "justified" or "j" => Drawing.TextAlignmentTypeValues.Justified,
                            "left" or "l" => Drawing.TextAlignmentTypeValues.Left,
                            _ => throw new ArgumentException($"Invalid align value: '{value}'. Valid values: left, center, right, justify.")
                        };
                    }
                    break;
                case "valign":
                {
                    var txBody = shape.TextBody;
                    var bodyPr = txBody?.GetFirstChild<Drawing.BodyProperties>();
                    if (bodyPr != null)
                    {
                        bodyPr.Anchor = value.ToLowerInvariant() switch
                        {
                            "top" or "t" => Drawing.TextAnchoringTypeValues.Top,
                            "center" or "ctr" or "middle" or "m" or "c" => Drawing.TextAnchoringTypeValues.Center,
                            "bottom" or "b" => Drawing.TextAnchoringTypeValues.Bottom,
                            _ => throw new ArgumentException($"Invalid valign value: '{value}'. Valid values: top, center, bottom.")
                        };
                    }
                    break;
                }
                case "gradientfill":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr != null)
                    {
                        spPr.RemoveAllChildren<Drawing.SolidFill>();
                        spPr.RemoveAllChildren<Drawing.NoFill>();
                        spPr.RemoveAllChildren<Drawing.GradientFill>();
                        // CONSISTENCY(shape-gradient-fill): reuse Add-branch parser.
                        spPr.AppendChild(BuildShapeGradientFill(value));
                    }
                    break;
                }
                case "line" or "border":
                {
                    // CONSISTENCY(shape-line): mirror Add — accept "none" or "color[:width[:style]]".
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) break;
                    spPr.RemoveAllChildren<Drawing.Outline>();
                    if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        spPr.AppendChild(new Drawing.Outline(new Drawing.NoFill()));
                        break;
                    }
                    var parts = value.Split(':');
                    var (lRgb, _) = ParseHelpers.SanitizeColorForOoxml(parts[0]);
                    var outline = new Drawing.Outline(
                        new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = lRgb }));
                    if (parts.Length > 1
                        && double.TryParse(parts[1], System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out var wpt))
                    {
                        outline.Width = (int)Math.Round(wpt * 12700);
                    }
                    if (parts.Length > 2)
                    {
                        var dash = parts[2].ToLowerInvariant() switch
                        {
                            "dash" => Drawing.PresetLineDashValues.Dash,
                            "dot" => Drawing.PresetLineDashValues.Dot,
                            "dashdot" => Drawing.PresetLineDashValues.DashDot,
                            "longdash" => Drawing.PresetLineDashValues.LargeDash,
                            "solid" => Drawing.PresetLineDashValues.Solid,
                            _ => (Drawing.PresetLineDashValues?)null
                        };
                        if (dash != null)
                            outline.AppendChild(new Drawing.PresetDash { Val = dash });
                    }
                    spPr.AppendChild(outline);
                    break;
                }
                case "alt" or "alttext" or "descr" or "description":
                {
                    var altNv = shape.NonVisualShapeProperties?
                        .GetFirstChild<XDR.NonVisualDrawingProperties>();
                    if (altNv != null) altNv.Description = value;
                    break;
                }
                default:
                    shpUnsupported.Add(key);
                    break;
            }
        }

        drawingsPart.WorksheetDrawing.Save();
        return shpUnsupported;
    }

    private List<string> SetSlicerByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var slIdx = int.Parse(m.Groups[1].Value);
        if (!TryFindSlicerByIndex(worksheet, slIdx, out var slicer, out _) || slicer == null)
            throw new ArgumentException($"slicer[{slIdx}] not found on sheet");

        var slicersPart = worksheet.GetPartsOfType<SlicersPart>().FirstOrDefault();
        var slUnsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "caption": slicer.Caption = value; break;
                case "style": slicer.Style = value; break;
                case "name": slicer.Name = value; break;
                case "rowheight":
                    if (uint.TryParse(value, out var rh)) slicer.RowHeight = rh;
                    else slUnsupported.Add(key);
                    break;
                case "columncount":
                    if (uint.TryParse(value, out var cc) && cc >= 1 && cc <= 20000)
                        slicer.ColumnCount = cc;
                    else slUnsupported.Add(key);
                    break;
                default: slUnsupported.Add(key); break;
            }
        }
        if (slicersPart?.Slicers != null) slicersPart.Slicers.Save(slicersPart);
        SaveWorksheet(worksheet);
        return slUnsupported;
    }

    // CONSISTENCY(table-column-path): mirror the col[M].prop= dotted form already
    // accepted on /Sheet/table[N] by exposing the column as a sub-path so users
    // can address it as a node and call Set with a flat property bag.
    private List<string> SetTableColumnByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var tIdx = int.Parse(m.Groups[1].Value);
        var cIdx = int.Parse(m.Groups[2].Value);
        var tParts = worksheet.TableDefinitionParts.ToList();
        if (tIdx < 1 || tIdx > tParts.Count)
            throw new ArgumentException($"Table index {tIdx} out of range (1..{tParts.Count})");
        var tbl = tParts[tIdx - 1].Table
            ?? throw new ArgumentException($"Table {tIdx} has no definition");
        var tCols = tbl.GetFirstChild<TableColumns>()?.Elements<TableColumn>().ToList();
        if (tCols == null || cIdx < 1 || cIdx > tCols.Count)
            throw new ArgumentException($"Column index {cIdx} out of range (1..{tCols?.Count ?? 0})");
        var tCol = tCols[cIdx - 1];
        var tcUnsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "name":
                {
                    tCol.Name = value;
                    // Sync the header-row cell so the worksheet matches the
                    // tableColumn @name. Excel rejects mismatch otherwise.
                    var refStr = tbl.Reference?.Value;
                    if (!string.IsNullOrEmpty(refStr) && (tbl.HeaderRowCount?.Value ?? 1) != 0)
                    {
                        var rParts = refStr.Split(':');
                        if (rParts.Length >= 1)
                        {
                            var (startCol, startRow) = ParseCellReference(rParts[0]);
                            var headerColIdx = ColumnNameToIndex(startCol) + (cIdx - 1);
                            var headerColLetter = IndexToColumnName(headerColIdx);
                            var headerCellRef = $"{headerColLetter}{startRow}";
                            var hdrWs = GetSheet(worksheet);
                            var hdrSheetData = hdrWs.GetFirstChild<SheetData>()
                                ?? hdrWs.AppendChild(new SheetData());
                            var hdrCell = FindOrCreateCell(hdrSheetData, headerCellRef);
                            hdrCell.CellValue = new CellValue(value);
                            hdrCell.DataType = CellValues.String;
                        }
                    }
                    break;
                }
                case "totalfunction" or "total":
                    tCol.TotalsRowFunction = value.ToLowerInvariant() switch
                    {
                        "sum" => TotalsRowFunctionValues.Sum,
                        "count" => TotalsRowFunctionValues.Count,
                        "average" or "avg" => TotalsRowFunctionValues.Average,
                        "max" => TotalsRowFunctionValues.Maximum,
                        "min" => TotalsRowFunctionValues.Minimum,
                        "stddev" => TotalsRowFunctionValues.StandardDeviation,
                        "var" => TotalsRowFunctionValues.Variance,
                        "countnums" => TotalsRowFunctionValues.CountNumbers,
                        "none" => TotalsRowFunctionValues.None,
                        "custom" => TotalsRowFunctionValues.Custom,
                        _ => throw new ArgumentException($"Invalid totalFunction: '{value}'.")
                    };
                    break;
                case "totallabel" or "label":
                    tCol.TotalsRowLabel = value;
                    break;
                case "formula":
                    tCol.CalculatedColumnFormula = new CalculatedColumnFormula(value);
                    break;
                default:
                    tcUnsupported.Add(key);
                    break;
            }
        }
        tParts[tIdx - 1].Table!.Save();
        SaveWorksheet(worksheet);
        return tcUnsupported;
    }

    private List<string> SetTableByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var tableIdx = int.Parse(m.Groups[1].Value);
        var tableParts = worksheet.TableDefinitionParts.ToList();
        if (tableIdx < 1 || tableIdx > tableParts.Count)
            throw new ArgumentException($"Table index {tableIdx} out of range (1..{tableParts.Count})");

        var table = tableParts[tableIdx - 1].Table
            ?? throw new ArgumentException($"Table {tableIdx} has no definition");

        var tblUnsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "name": table.Name = value; break;
                case "displayname": table.DisplayName = value; break;
                case "headerrow": table.HeaderRowCount = IsTruthy(value) ? 1u : 0u; break;
                case "totalrow":
                case "showtotals":
                    var totalRowEnabled = IsTruthy(value);
                    table.TotalsRowShown = totalRowEnabled;
                    table.TotalsRowCount = totalRowEnabled ? 1u : 0u;
                    break;
                case "style":
                    var styleInfo = table.GetFirstChild<TableStyleInfo>();
                    if (styleInfo != null) styleInfo.Name = value;
                    else table.AppendChild(new TableStyleInfo
                    {
                        Name = value, ShowFirstColumn = false, ShowLastColumn = false,
                        ShowRowStripes = true, ShowColumnStripes = false
                    });
                    break;
                case "ref":
                {
                    var newRef = value.ToUpperInvariant();
                    // Grow/shrink <x:tableColumns> to match the new column count.
                    // Excel rejects the file when tableColumns.Count mismatches the
                    // ref width. On grow, append default ColumnN entries; on shrink,
                    // trim trailing entries.
                    var newParts = newRef.Split(':');
                    if (newParts.Length == 2)
                    {
                        var (nsc, _) = ParseCellReference(newParts[0]);
                        var (nec, _) = ParseCellReference(newParts[1]);
                        int newColCount = ColumnNameToIndex(nec) - ColumnNameToIndex(nsc) + 1;
                        var tc = table.GetFirstChild<TableColumns>();
                        if (tc != null && newColCount > 0)
                        {
                            var cols = tc.Elements<TableColumn>().ToList();
                            if (newColCount > cols.Count)
                            {
                                var existingIds = cols.Select(c => c.Id?.Value ?? 0u).ToList();
                                var existingNames = new HashSet<string>(
                                    cols.Select(c => c.Name?.Value ?? string.Empty),
                                    StringComparer.OrdinalIgnoreCase);
                                uint nextId = existingIds.Count > 0 ? existingIds.Max() + 1 : 1u;
                                for (int i = cols.Count; i < newColCount; i++)
                                {
                                    var baseName = $"Column{i + 1}";
                                    var name = baseName;
                                    int dedup = 2;
                                    while (!existingNames.Add(name))
                                        name = $"{baseName}{dedup++}";
                                    tc.AppendChild(new TableColumn { Id = nextId++, Name = name });
                                }
                            }
                            else if (newColCount < cols.Count)
                            {
                                for (int i = cols.Count - 1; i >= newColCount; i--)
                                    cols[i].Remove();
                            }
                            tc.Count = (uint)newColCount;
                        }
                    }
                    table.Reference = newRef;
                    var af = table.GetFirstChild<AutoFilter>();
                    if (af != null) af.Reference = newRef;
                    break;
                }
                case "showrowstripes" or "bandedrows" or "bandrows":
                {
                    var si = table.GetFirstChild<TableStyleInfo>();
                    if (si != null) si.ShowRowStripes = IsTruthy(value);
                    break;
                }
                case "showcolstripes" or "showcolumnstripes" or "bandedcols" or "bandcols":
                {
                    var si = table.GetFirstChild<TableStyleInfo>();
                    if (si != null) si.ShowColumnStripes = IsTruthy(value);
                    break;
                }
                case "showfirstcolumn" or "firstcol" or "firstcolumn":
                {
                    var si = table.GetFirstChild<TableStyleInfo>();
                    if (si != null) si.ShowFirstColumn = IsTruthy(value);
                    break;
                }
                case "showlastcolumn" or "lastcol" or "lastcolumn":
                {
                    var si = table.GetFirstChild<TableStyleInfo>();
                    if (si != null) si.ShowLastColumn = IsTruthy(value);
                    break;
                }
                case var k when k.StartsWith("col[") || k.StartsWith("column["):
                {
                    var tblColMatch = Regex.Match(k, @"^col(?:umn)?\[(\d+)\]\.(.+)$", RegexOptions.IgnoreCase);
                    if (!tblColMatch.Success) { tblUnsupported.Add(key); break; }
                    var colIdx = int.Parse(tblColMatch.Groups[1].Value);
                    var colProp = tblColMatch.Groups[2].Value.ToLowerInvariant();
                    var tableCols = table.GetFirstChild<TableColumns>()?.Elements<TableColumn>().ToList();
                    if (tableCols == null || colIdx < 1 || colIdx > tableCols.Count)
                        throw new ArgumentException($"Column index {colIdx} out of range (1..{tableCols?.Count ?? 0})");
                    var col = tableCols[colIdx - 1];
                    switch (colProp)
                    {
                        case "name": col.Name = value; break;
                        case "totalfunction" or "total":
                            col.TotalsRowFunction = value.ToLowerInvariant() switch
                            {
                                "sum" => TotalsRowFunctionValues.Sum,
                                "count" => TotalsRowFunctionValues.Count,
                                "average" or "avg" => TotalsRowFunctionValues.Average,
                                "max" => TotalsRowFunctionValues.Maximum,
                                "min" => TotalsRowFunctionValues.Minimum,
                                "stddev" => TotalsRowFunctionValues.StandardDeviation,
                                "var" => TotalsRowFunctionValues.Variance,
                                "countnums" => TotalsRowFunctionValues.CountNumbers,
                                "none" => TotalsRowFunctionValues.None,
                                "custom" => TotalsRowFunctionValues.Custom,
                                _ => throw new ArgumentException($"Invalid totalFunction: '{value}'. Valid: sum, count, average, max, min, stddev, var, countNums, none, custom.")
                            };
                            break;
                        case "totallabel" or "label":
                            col.TotalsRowLabel = value;
                            break;
                        case "formula":
                            col.CalculatedColumnFormula = new CalculatedColumnFormula(value);
                            break;
                        default: tblUnsupported.Add(key); break;
                    }
                    break;
                }
                default: tblUnsupported.Add(key); break;
            }
        }

        tableParts[tableIdx - 1].Table!.Save();
        return tblUnsupported;
    }

    private List<string> SetCommentByPath(Match m, WorksheetPart worksheet, string sheetName, Dictionary<string, string> properties)
    {
        var cmtIndex = int.Parse(m.Groups[1].Value);
        var commentsPart = worksheet.WorksheetCommentsPart;
        if (commentsPart?.Comments == null)
            throw new ArgumentException($"No comments found in sheet: {sheetName}");

        var cmtList = commentsPart.Comments.GetFirstChild<CommentList>();
        var cmtElement = cmtList?.Elements<Comment>().ElementAtOrDefault(cmtIndex - 1)
            ?? throw new ArgumentException($"Comment [{cmtIndex}] not found");

        var cmtUnsupported = new List<string>();
        // CONSISTENCY(xlsx/comment-font): C8 — font.* props on Set rewrite the
        // single <x:r><x:rPr>, reusing BuildCommentRunProperties. When `text` and
        // `font.*` appear together, text wins the run payload and font.* supplies
        // the rPr. When only font.* appears (no text), preserve the existing run
        // text and just rebuild rPr.
        string? newCmtText = properties.TryGetValue("text", out var tVal) ? tVal : null;
        bool hasFontProp = properties.Keys.Any(k =>
            k.StartsWith("font.", StringComparison.OrdinalIgnoreCase));
        if (newCmtText != null || hasFontProp)
        {
            string runText = newCmtText
                ?? string.Concat(cmtElement.CommentText?.Elements<Run>()
                    .SelectMany(r => r.Elements<Text>()).Select(t => t.Text)
                    ?? Array.Empty<string>());
            cmtElement.CommentText = new CommentText(
                new Run(
                    BuildCommentRunProperties(properties),
                    new Text(runText) { Space = SpaceProcessingModeValues.Preserve }
                )
            );
        }
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "text":
                case var k1 when k1.StartsWith("font."):
                    break;
                case "ref":
                    cmtElement.Reference = value.ToUpperInvariant();
                    break;
                case "author":
                    var authors = commentsPart.Comments.GetFirstChild<Authors>()!;
                    var existingAuthors = authors.Elements<Author>().ToList();
                    var aIdx = existingAuthors.FindIndex(a => a.Text == value);
                    if (aIdx >= 0)
                        cmtElement.AuthorId = (uint)aIdx;
                    else
                    {
                        authors.AppendChild(new Author(value));
                        cmtElement.AuthorId = (uint)existingAuthors.Count;
                    }
                    break;
                default:
                    cmtUnsupported.Add(key);
                    break;
            }
        }

        commentsPart.Comments.Save();
        return cmtUnsupported;
    }

    private List<string> SetCfByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var cfIdx = int.Parse(m.Groups[1].Value);
        var ws = GetSheet(worksheet);
        var cfElements = ws.Elements<ConditionalFormatting>().ToList();
        if (cfIdx < 1 || cfIdx > cfElements.Count)
            throw new ArgumentException($"CF {cfIdx} not found (total: {cfElements.Count})");

        var cf = cfElements[cfIdx - 1];
        var unsup = new List<string>();
        var rule = cf.Elements<ConditionalFormattingRule>().FirstOrDefault();

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "sqref":
                    cf.SequenceOfReferences = new ListValue<StringValue>(
                        value.Split(' ').Select(s => new StringValue(s)));
                    break;
                case "color":
                    var dbColor = rule?.GetFirstChild<DataBar>()?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>();
                    if (dbColor != null) { dbColor.Rgb = ParseHelpers.NormalizeArgbColor(value); }
                    else unsup.Add(key);
                    break;
                case "mincolor":
                    var csColors = rule?.GetFirstChild<ColorScale>()?.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToList();
                    if (csColors != null && csColors.Count >= 2)
                    { csColors[0].Rgb = ParseHelpers.NormalizeArgbColor(value); }
                    else unsup.Add(key);
                    break;
                case "maxcolor":
                    var csColors2 = rule?.GetFirstChild<ColorScale>()?.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToList();
                    if (csColors2 != null && csColors2.Count >= 2)
                    { csColors2[^1].Rgb = ParseHelpers.NormalizeArgbColor(value); }
                    else unsup.Add(key);
                    break;
                case "iconset":
                case "icons":
                    var iconSetEl = rule?.GetFirstChild<IconSet>();
                    if (iconSetEl != null)
                        iconSetEl.IconSetValue = new EnumValue<IconSetValues>(ParseIconSetValues(value));
                    else unsup.Add(key);
                    break;
                case "reverse":
                    var isEl = rule?.GetFirstChild<IconSet>();
                    if (isEl != null) isEl.Reverse = IsTruthy(value);
                    else unsup.Add(key);
                    break;
                case "showvalue":
                    var isEl2 = rule?.GetFirstChild<IconSet>();
                    if (isEl2 != null) isEl2.ShowValue = IsTruthy(value);
                    else unsup.Add(key);
                    break;
                default:
                    unsup.Add(key);
                    break;
            }
        }
        SaveWorksheet(worksheet);
        return unsup;
    }

    private List<string> SetChartAxisByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var caChartIdx = int.Parse(m.Groups[1].Value);
        var caRole = m.Groups[2].Value;
        var caDrawingsPart = worksheet.DrawingsPart
            ?? throw new ArgumentException("No charts in this sheet");
        var caAllCharts = GetExcelCharts(caDrawingsPart);
        if (caChartIdx < 1 || caChartIdx > caAllCharts.Count)
            throw new ArgumentException($"Chart {caChartIdx} not found (total: {caAllCharts.Count})");
        var caChartInfo = caAllCharts[caChartIdx - 1];
        if (caChartInfo.IsExtended || caChartInfo.StandardPart == null)
            throw new ArgumentException("Axis Set not supported on extended charts.");
        var axUnsupported = ChartHelper.SetAxisProperties(
            caChartInfo.StandardPart, caRole, properties);
        caChartInfo.StandardPart.ChartSpace?.Save();
        return axUnsupported;
    }

    private List<string> SetChartByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var chartIdx = int.Parse(m.Groups[1].Value);
        var drawingsPart = worksheet.DrawingsPart
            ?? throw new ArgumentException("No charts in this sheet");
        var excelCharts = GetExcelCharts(drawingsPart);
        if (chartIdx < 1 || chartIdx > excelCharts.Count)
            throw new ArgumentException($"Chart {chartIdx} not found (total: {excelCharts.Count})");
        var chartInfo = excelCharts[chartIdx - 1];

        // If series sub-path, prefix all properties with series{N}. for ChartSetter
        var chartProps = properties;
        var isSeriesPath = m.Groups[2].Success;
        if (isSeriesPath)
        {
            var seriesIdx = int.Parse(m.Groups[2].Value);
            chartProps = new Dictionary<string, string>();
            foreach (var (key, value) in properties)
                chartProps[$"series{seriesIdx}.{key}"] = value;
        }

        // Chart-level position/size Set — TwoCellAnchor mutation. Skip for series
        // sub-paths (series don't have their own position). Accepts x/y/width/height
        // in the same units as OLE Set and chart Add.
        // CONSISTENCY(chart-position-set): mirrors PPTX path so users learn one
        // vocabulary for all three doc types. Excel mutates a TwoCellAnchor instead
        // of a GraphicFrame Transform because xlsx charts are cell-anchored.
        if (!isSeriesPath)
        {
            var positionUnsupported = ApplyChartPositionSet(
                drawingsPart, chartIdx, chartProps);
            foreach (var k in new[] { "x", "y", "width", "height" })
            {
                var matched = chartProps.Keys
                    .FirstOrDefault(key => key.Equals(k, StringComparison.OrdinalIgnoreCase));
                if (matched != null && !positionUnsupported.Contains(matched))
                    chartProps.Remove(matched);
            }
        }

        if (chartInfo.StandardPart != null)
        {
            var unsup = ChartHelper.SetChartProperties(chartInfo.StandardPart, chartProps);
            chartInfo.StandardPart.ChartSpace?.Save();
            return unsup;
        }
        else if (chartInfo.ExtendedPart != null)
        {
            // cx:chart — delegates to ChartExBuilder.SetChartProperties.
            return ChartExBuilder.SetChartProperties(chartInfo.ExtendedPart, chartProps);
        }
        else
        {
            return chartProps.Keys.ToList();
        }
    }

    private List<string> SetPivotTableByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var ptIdx = int.Parse(m.Groups[1].Value);
        var pivotParts = worksheet.PivotTableParts.ToList();
        if (ptIdx < 1 || ptIdx > pivotParts.Count)
            throw new ArgumentException($"PivotTable {ptIdx} not found");
        return PivotTableHelper.SetPivotTableProperties(pivotParts[ptIdx - 1], properties);
    }

    private List<string> SetCellRunByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var runCellRef = m.Groups[1].Value.ToUpperInvariant();
        var runIdx = int.Parse(m.Groups[2].Value);

        var runSheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet data not found");
        var runCell = FindOrCreateCell(runSheetData, runCellRef);

        if (runCell.DataType?.Value != CellValues.SharedString ||
            !int.TryParse(runCell.CellValue?.Text, out var sstIdx))
            throw new ArgumentException($"Cell {runCellRef} is not a rich text cell");

        var sstPart = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        var ssi = sstPart?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(sstIdx)
            ?? throw new ArgumentException($"SharedString entry {sstIdx} not found");

        var runs = ssi.Elements<Run>().ToList();
        if (runIdx < 1 || runIdx > runs.Count)
            throw new ArgumentException($"Run index {runIdx} out of range (1-{runs.Count})");

        var run = runs[runIdx - 1];
        var rProps = run.RunProperties ?? run.PrependChild(new RunProperties());

        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "text" or "value":
                    var textEl = run.GetFirstChild<Text>();
                    if (textEl != null) textEl.Text = value;
                    else run.AppendChild(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
                    break;
                case "bold":
                    rProps.RemoveAllChildren<Bold>();
                    if (ParseHelpers.IsTruthy(value)) rProps.InsertAt(new Bold(), 0);
                    break;
                case "italic":
                    rProps.RemoveAllChildren<Italic>();
                    if (ParseHelpers.IsTruthy(value)) rProps.AppendChild(new Italic());
                    break;
                case "strike":
                    rProps.RemoveAllChildren<Strike>();
                    if (ParseHelpers.IsTruthy(value)) rProps.AppendChild(new Strike());
                    break;
                case "underline":
                    rProps.RemoveAllChildren<Underline>();
                    if (!string.IsNullOrEmpty(value) && value != "false" && value != "none")
                    {
                        var ul = new Underline();
                        if (value.ToLowerInvariant() == "double") ul.Val = UnderlineValues.Double;
                        rProps.AppendChild(ul);
                    }
                    break;
                case "superscript":
                    rProps.RemoveAllChildren<VerticalTextAlignment>();
                    if (ParseHelpers.IsTruthy(value))
                        rProps.AppendChild(new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Superscript });
                    break;
                case "subscript":
                    rProps.RemoveAllChildren<VerticalTextAlignment>();
                    if (ParseHelpers.IsTruthy(value))
                        rProps.AppendChild(new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Subscript });
                    break;
                case "size":
                    rProps.RemoveAllChildren<FontSize>();
                    rProps.AppendChild(new FontSize { Val = ParseHelpers.ParseFontSize(value) });
                    break;
                case "color":
                    rProps.RemoveAllChildren<Color>();
                    rProps.AppendChild(new Color { Rgb = ParseHelpers.NormalizeArgbColor(value) });
                    break;
                case "font":
                    rProps.RemoveAllChildren<RunFont>();
                    rProps.AppendChild(new RunFont { Val = value });
                    break;
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        ReorderRunProperties(rProps);
        sstPart!.SharedStringTable!.Save();
        SaveWorksheet(worksheet);
        return unsupported;
    }
}
