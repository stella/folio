using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlProjector;

internal static class Program
{
    private static int Main(string[] args)
    {
        if (args.Length != 1)
        {
            Console.Error.WriteLine("usage: OpenXmlProjector <docx-path>");
            return 2;
        }

        try
        {
            var projection = StructuralProjector.Project(args[0]);
            var json = JsonSerializer.Serialize(
                projection,
                JsonOptions.Default
            );
            Console.WriteLine(json);
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error.Message);
            return 2;
        }
    }
}

internal static class JsonOptions
{
    public static readonly JsonSerializerOptions Default = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = true,
    };
}

internal sealed record SdtProjection
{
    public required string Scope { get; init; }
    public required string SdtType { get; init; }
    public string? Alias { get; init; }
    public string? Tag { get; init; }
    public string? Lock { get; init; }
    public required int ChildCount { get; init; }
}

internal sealed record StructuralProjection
{
    public required int SchemaVersion { get; init; }
    public required int TotalParagraphs { get; init; }
    public required int TotalTables { get; init; }
    public required int TopLevelBlocks { get; init; }
    public required List<SdtProjection> Sdts { get; init; }
    public required Dictionary<string, int> SdtCountsByType { get; init; }
}

internal static class StructuralProjector
{
    private const int SchemaVersion = 1;

    private const string W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    private static readonly (string LocalName, string NamespaceUri, string Normalised)[] SdtTypeElements =
    [
        ("richText", W, "richText"),
        ("text", W, "plainText"),
        ("date", W, "date"),
        ("dropDownList", W, "dropdown"),
        ("comboBox", W, "comboBox"),
        ("checkbox", W, "checkbox"),
        ("picture", W, "picture"),
        ("docPartObj", W, "buildingBlockGallery"),
        ("docPartList", W, "buildingBlockGallery"),
        ("group", W, "group"),
    ];

    private static readonly HashSet<string> LockValues =
    [
        "sdtLocked",
        "contentLocked",
        "sdtContentLocked",
        "unlocked",
    ];

    private static readonly HashSet<(string LocalName, string NamespaceUri)> InlineTags =
    [
        ("r", W),
        ("hyperlink", W),
        ("fldSimple", W),
        ("sdt", W),
        ("oMath", M),
        ("oMathPara", M),
    ];

    private static readonly HashSet<(string LocalName, string NamespaceUri)> InlineParentTags =
    [
        ("p", W),
        ("hyperlink", W),
        ("smartTag", W),
    ];

    private static readonly HashSet<(string LocalName, string NamespaceUri)> ScopeTransparentTags =
    [
        ("sdtContent", W),
        ("sdt", W),
        ("ins", W),
        ("del", W),
        ("moveFrom", W),
        ("moveTo", W),
    ];

    private static readonly HashSet<(string LocalName, string NamespaceUri)> TopLevelBlockTags =
    [
        ("p", W),
        ("tbl", W),
        ("sdt", W),
    ];

    public static StructuralProjection Project(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        var body = document.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("document has no body");

        var topLevelBlocks = body
            .ChildElements.Count(child => Matches(child, TopLevelBlockTags));

        var totalParagraphs = body
            .Descendants<Paragraph>()
            .Count(paragraph => !InTextbox(paragraph));

        var totalTables = body
            .Descendants<Table>()
            .Count(table => !InTextbox(table));

        var sdts = body
            .Descendants()
            .Where(element => IsTag(element, "sdt", W) && !InTextbox(element))
            .Select(ProjectSdt)
            .ToList();

        var counts = new Dictionary<string, int>();
        foreach (var sdt in sdts)
        {
            counts[sdt.SdtType] = counts.GetValueOrDefault(sdt.SdtType) + 1;
            counts[sdt.Scope] = counts.GetValueOrDefault(sdt.Scope) + 1;
        }

        return new StructuralProjection
        {
            SchemaVersion = SchemaVersion,
            TotalParagraphs = totalParagraphs,
            TotalTables = totalTables,
            TopLevelBlocks = topLevelBlocks,
            Sdts = sdts,
            SdtCountsByType = counts,
        };
    }

    private static SdtProjection ProjectSdt(OpenXmlElement sdtElement)
    {
        var sdtPr = sdtElement.ChildElements.FirstOrDefault(child => IsTag(child, "sdtPr", W));
        var sdtContent = sdtElement.ChildElements.FirstOrDefault(child =>
            IsTag(child, "sdtContent", W)
        );

        var scopeParent = sdtElement.Parent;
        while (scopeParent is not null && Matches(scopeParent, ScopeTransparentTags))
        {
            scopeParent = scopeParent.Parent;
        }

        var scope = scopeParent is not null && Matches(scopeParent, InlineParentTags)
            ? "inline"
            : "block";

        var alias = TextAttr(sdtPr, "alias");
        var tag = TextAttr(sdtPr, "tag");
        var lockValue = TextAttr(sdtPr, "lock");
        if (lockValue is not null && !LockValues.Contains(lockValue))
        {
            lockValue = null;
        }

        var childCount = 0;
        if (sdtContent is not null)
        {
            childCount = scope == "block"
                ? sdtContent.ChildElements.Count(child =>
                    IsTag(child, "p", W) || IsTag(child, "tbl", W) || IsTag(child, "sdt", W)
                )
                : CountInlineChildren(sdtContent);
        }

        return new SdtProjection
        {
            Scope = scope,
            SdtType = DetectSdtType(sdtPr),
            Alias = alias,
            Tag = tag,
            Lock = lockValue,
            ChildCount = childCount,
        };
    }

    private static string DetectSdtType(OpenXmlElement? sdtPr)
    {
        if (sdtPr is null)
        {
            return "richText";
        }

        foreach (var (localName, namespaceUri, normalised) in SdtTypeElements)
        {
            if (sdtPr.ChildElements.Any(child => IsTag(child, localName, namespaceUri)))
            {
                return normalised;
            }
        }

        if (sdtPr.ChildElements.Any(child => IsTag(child, "checkbox", W14)))
        {
            return "checkbox";
        }

        return "richText";
    }

    private static string? TextAttr(OpenXmlElement? parent, string tagLocalName)
    {
        if (parent is null)
        {
            return null;
        }

        var child = parent.ChildElements.FirstOrDefault(element => IsTag(element, tagLocalName, W));
        if (child is null)
        {
            return null;
        }

        return child.GetAttributes()
            .FirstOrDefault(attribute =>
                attribute.LocalName == "val"
                && attribute.NamespaceUri == W
            )
            .Value;
    }

    private static int CountInlineChildren(OpenXmlElement sdtContent)
    {
        var count = 0;
        var fieldDepth = 0;

        foreach (var child in sdtContent.ChildElements)
        {
            if (!IsTag(child, "r", W))
            {
                if (fieldDepth == 0 && Matches(child, InlineTags))
                {
                    count += 1;
                }
                continue;
            }

            var charType = RunFieldCharType(child);
            if (charType == "begin")
            {
                if (fieldDepth == 0)
                {
                    count += 1;
                }
                fieldDepth += 1;
                continue;
            }

            if (charType == "end")
            {
                if (fieldDepth > 0)
                {
                    fieldDepth -= 1;
                }
                continue;
            }

            if (fieldDepth == 0)
            {
                count += 1;
            }
        }

        return count;
    }

    private static string? RunFieldCharType(OpenXmlElement runElement)
    {
        var fieldChar = runElement.ChildElements.FirstOrDefault(child =>
            IsTag(child, "fldChar", W)
        );
        if (fieldChar is null)
        {
            return null;
        }

        return fieldChar.GetAttributes()
            .FirstOrDefault(attribute =>
                attribute.LocalName == "fldCharType" && attribute.NamespaceUri == W
            )
            .Value;
    }

    private static bool InTextbox(OpenXmlElement element)
    {
        var parent = element.Parent;
        while (parent is not null)
        {
            if (IsTag(parent, "txbxContent", W))
            {
                return true;
            }
            parent = parent.Parent;
        }
        return false;
    }

    private static bool IsTag(OpenXmlElement element, string localName, string namespaceUri) =>
        element.LocalName == localName && element.NamespaceUri == namespaceUri;

    private static bool Matches(
        OpenXmlElement element,
        IEnumerable<(string LocalName, string NamespaceUri)> tags
    ) => tags.Any(tag => IsTag(element, tag.LocalName, tag.NamespaceUri));
}
