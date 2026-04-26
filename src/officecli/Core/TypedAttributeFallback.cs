// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;

namespace OfficeCli.Core;

/// <summary>
/// Generic dotted-key fallback for setting an OOXML attribute on a child
/// element of a known parent container. Sister to
/// <see cref="GenericXmlQuery.TryCreateTypedChild"/>, which only covers
/// "single val" leaf elements.
///
/// <para>
/// The shape it accepts is <c>elementLocalName.attrLocalName=value</c>.
/// For example, <c>ind.firstLine=240</c> resolves to
/// <c>&lt;w:ind w:firstLine="240"/&gt;</c> under the parent. If the child
/// element already exists, the attribute is merged in (the helper preserves
/// other attrs the caller did not pass) — so a chain of
/// <c>set ind.left=720</c> followed by <c>set ind.firstLine=240</c>
/// produces a single <c>&lt;w:ind/&gt;</c> with both attrs, not two
/// elements or one overwrite.
/// </para>
///
/// <para>
/// Validation is delegated to the OpenXML SDK: we round-trip the requested
/// element through <c>InnerXml</c>, and reject anything the SDK parses as
/// <see cref="OpenXmlUnknownElement"/> or whose attribute did not bind. This
/// is the same trick <c>TryCreateTypedChild</c> uses, so the schema rules
/// are identical: known element + known attr only, no garbage XML.
/// </para>
///
/// <para>
/// Aliases: a small map normalizes user-facing names (<c>font</c>,
/// <c>shading</c>, <c>underline</c>, <c>border</c>) to the OOXML local
/// names (<c>rFonts</c>, <c>shd</c>, <c>u</c>, <c>pBdr</c>) so the fallback
/// stays consistent with the curated vocabulary in the rest of the
/// handler.
/// </para>
/// </summary>
internal static class TypedAttributeFallback
{
    /// <summary>
    /// User-facing element-name aliases. Keep this small and aligned with
    /// the curated vocabulary used elsewhere in the Word handler. Adding an
    /// alias here also implicitly extends what the dotted fallback accepts.
    /// </summary>
    private static readonly Dictionary<string, string> ElementAliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["font"]      = "rFonts",
        ["shading"]   = "shd",
        ["underline"] = "u",
        ["border"]    = "pBdr",
    };

    /// <summary>
    /// Attempt to set <paramref name="value"/> as an attribute on a child
    /// element of <paramref name="parent"/>, where the dotted key has the
    /// form <c>"elementName.attrName"</c>. Returns false (and does not
    /// modify <paramref name="parent"/>) if the dotted shape is not
    /// recognized by the SDK as a valid element/attr pair under this
    /// parent.
    /// </summary>
    public static bool TrySet(OpenXmlElement parent, string dottedKey, string value)
    {
        var dot = dottedKey.IndexOf('.');
        if (dot <= 0 || dot == dottedKey.Length - 1) return false;
        var elementLocal = dottedKey[..dot];
        var attrLocal    = dottedKey[(dot + 1)..];
        if (ElementAliases.TryGetValue(elementLocal, out var aliased))
            elementLocal = aliased;

        var nsUri  = parent.NamespaceUri;
        var prefix = parent.Prefix;
        // Detached probe elements (e.g. `new StyleParagraphProperties()` not
        // yet attached to a part) report empty Prefix / NamespaceUri. Fall
        // back to the Word namespace — this fallback is currently only wired
        // into the Word handler. If/when reused for PPTX/XLSX, route the
        // namespace through the caller instead of hardcoding here.
        if (string.IsNullOrEmpty(nsUri) || string.IsNullOrEmpty(prefix))
        {
            nsUri  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            prefix = "w";
        }

        // Validate (element, attr) is a known SDK pair under this parent by
        // round-tripping through InnerXml. If SDK does not recognize either
        // side, the parsed result is OpenXmlUnknownElement — reject so we
        // never write garbage XML. This is the same approach
        // TryCreateTypedChild uses for single-val leaf elements.
        OpenXmlElement sample;
        try
        {
            var escapedVal = System.Security.SecurityElement.Escape(value);
            var temp = parent.CloneNode(false);
            temp.InnerXml = $"<{prefix}:{elementLocal} xmlns:{prefix}=\"{nsUri}\" {prefix}:{attrLocal}=\"{escapedVal}\"/>";
            // Clone (true) detaches the parsed element from its temporary
            // parent so it can be appended into the real tree later. Without
            // this, AppendChild throws "already part of a tree".
            var first = temp.FirstChild?.CloneNode(true);
            if (first is null or OpenXmlUnknownElement) return false;
            sample = (OpenXmlElement)first;
        }
        catch
        {
            return false;
        }

        // Validation: any typed attribute that survived parsing means the
        // (element, attr) pair was recognized by the SDK. If the user's
        // attr landed in ExtendedAttributes instead, the schema doesn't
        // know it (typo case like `ind.notAnAttr`) — reject.
        //
        // Note: SDK normalizes some legacy attr names (e.g. `w:left` →
        // `w:start` for bidi-aware indentation). We trust that
        // normalization rather than insisting the typed attr's local name
        // exactly match the user's input — both forms are schema-valid;
        // the SDK's canonical form is what gets written.
        if (sample.ExtendedAttributes.Any())
            return false;
        if (!sample.GetAttributes().Any())
            return false;

        // Apply: merge into existing child if present (copy each typed attr
        // from the sample so SDK normalization is preserved); otherwise
        // attach the sample as a new child. AppendChild is used rather than
        // AddChild because the latter can refuse schema-valid children when
        // the parent is a fresh detached probe with no document context —
        // the round-trip parse above already validated the pair.
        var existing = parent.ChildElements.FirstOrDefault(e =>
            e.LocalName.Equals(elementLocal, StringComparison.OrdinalIgnoreCase));
        if (existing != null)
        {
            foreach (var a in sample.GetAttributes())
                existing.SetAttribute(a);
            return true;
        }

        parent.AppendChild(sample);
        return true;
    }
}
