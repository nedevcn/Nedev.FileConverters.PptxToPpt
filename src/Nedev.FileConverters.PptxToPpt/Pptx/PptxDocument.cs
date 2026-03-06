using System.IO.Compression;
using System.Xml.Linq;

namespace Nedev.FileConverters.PptxToPpt.Pptx;

public sealed class PptxDocument : IDisposable
{
    private static readonly XNamespace RelsNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace OfficeNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private readonly ZipArchive _archive;
    private readonly Dictionary<string, byte[]> _cachedFiles = new();
    private bool _disposed;

    public PptxDocument(string filePath)
    {
        var stream = File.OpenRead(filePath);
        _archive = new ZipArchive(stream, ZipArchiveMode.Read);
    }

    public PptxDocument(Stream stream)
    {
        _archive = new ZipArchive(stream, ZipArchiveMode.Read);
    }

    public string[] GetAllFiles()
    {
        return _archive.Entries.Select(e => e.FullName).ToArray();
    }

    public byte[] GetFile(string path)
    {
        if (_cachedFiles.TryGetValue(path, out var cached))
            return cached;

        var entry = _archive.GetEntry(path);
        if (entry == null)
            return Array.Empty<byte>();

        using var ms = new MemoryStream();
        entry.Open().CopyTo(ms);
        var data = ms.ToArray();
        _cachedFiles[path] = data;
        return data;
    }

    public XDocument GetXml(string path)
    {
        var data = GetFile(path);
        if (data.Length == 0)
            return new XDocument();

        using var ms = new MemoryStream(data);
        return XDocument.Load(ms);
    }

    public IEnumerable<string> GetRelationships(string basePath)
    {
        var relPath = basePath.EndsWith("/") 
            ? basePath + "_rels/.rels" 
            : Path.GetDirectoryName(basePath)?.Replace('\\', '/') + "/_rels/" + Path.GetFileName(basePath) + ".rels";

        if (string.IsNullOrEmpty(relPath) || relPath == "/_rels/.rels")
            relPath = "_rels/.rels";

        var relDoc = GetXml(relPath);
        if (relDoc.Root == null)
            return Enumerable.Empty<string>();

        return relDoc.Root.Elements(RelsNs + "Relationship")
            .Select(r => r.Attribute("Target")?.Value)
            .Where(t => !string.IsNullOrEmpty(t))
            .Select(t => ResolveRelativePath(basePath, t!));
    }

    private static string ResolveRelativePath(string basePath, string target)
    {
        if (target.StartsWith("/"))
            return target.TrimStart('/');

        var dir = Path.GetDirectoryName(basePath)?.Replace('\\', '/') ?? "";
        if (dir.EndsWith("/"))
            return dir + target;
        return dir + "/" + target;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _archive.Dispose();
    }
}

public sealed class PptxSlide
{
    public int Index { get; set; }
    public string Path { get; set; } = "";
    public XDocument? Xml { get; set; }
    public Dictionary<int, XDocument> NotesXml { get; } = new();
}

public sealed class PptxPresentation
{
    public string RootPath { get; set; } = "ppt/presentation.xml";
    public XDocument? Xml { get; set; }
    public List<PptxSlide> Slides { get; } = new();
    public Dictionary<int, XDocument> SlideLayouts { get; } = new();
    public Dictionary<int, XDocument> SlideMasters { get; } = new();
    public Dictionary<string, byte[]> MediaFiles { get; } = new();
    public Dictionary<string, XDocument> ThemeDocuments { get; } = new();
    public Dictionary<string, int> Fonts { get; } = new();
    public XDocument? MainMaster { get; set; }
}

public sealed class PptxParser
{
    private static readonly XNamespace OfficeNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public async Task<PptxPresentation> ParseAsync(string filePath)
    {
        return await Task.Run(() => Parse(filePath));
    }

    public PptxPresentation Parse(string filePath)
    {
        using var doc = new PptxDocument(filePath);
        return ParseDocument(doc);
    }

    public PptxPresentation Parse(Stream stream)
    {
        using var doc = new PptxDocument(stream);
        return ParseDocument(doc);
    }

    private PptxPresentation ParseDocument(PptxDocument doc)
    {
        var presentation = new PptxPresentation();

        presentation.Xml = doc.GetXml("ppt/presentation.xml");
        if (presentation.Xml.Root == null)
            throw new InvalidOperationException("Invalid PPTX: presentation.xml not found or empty");

        var ns = presentation.Xml.Root.GetDefaultNamespace();

        var slideIds = presentation.Xml.Root.Element(ns + "sldIdLst")?.Elements(ns + "sldId") ?? Enumerable.Empty<XElement>();
        int index = 0;
        foreach (var slideId in slideIds)
        {
            var rId = slideId.Attribute(OfficeNs + "id")?.Value;
            if (string.IsNullOrEmpty(rId)) continue;

            var slidePath = ResolveRelationship(doc, "ppt/presentation.xml", rId);
            if (string.IsNullOrEmpty(slidePath)) continue;

            var slide = new PptxSlide
            {
                Index = index++,
                Path = slidePath,
                Xml = doc.GetXml(slidePath)
            };

            var notesPath = slidePath.Replace("slides/slide", "notesSlides/notesSlide");
            var notesXml = doc.GetXml(notesPath);
            if (notesXml.Root != null)
            {
                slide.NotesXml[slide.Index] = notesXml;
            }

            presentation.Slides.Add(slide);
        }

        var layoutIds = presentation.Xml.Root.Element(ns + "sldLayoutIdLst")?.Elements(ns + "sldLayoutId") ?? Enumerable.Empty<XElement>();
        foreach (var layoutId in layoutIds)
        {
            var rId = layoutId.Attribute(OfficeNs + "id")?.Value;
            if (string.IsNullOrEmpty(rId)) continue;

            var layoutPath = ResolveRelationship(doc, "ppt/presentation.xml", rId);
            if (string.IsNullOrEmpty(layoutPath)) continue;

            int layoutIndex = presentation.SlideLayouts.Count;
            var layoutXml = doc.GetXml(layoutPath);
            if (layoutXml.Root != null)
            {
                presentation.SlideLayouts[layoutIndex] = layoutXml;
            }
        }

        var masterIds = presentation.Xml.Root.Element(ns + "sldMasterIdLst")?.Elements(ns + "sldMasterId") ?? Enumerable.Empty<XElement>();
        foreach (var masterId in masterIds)
        {
            var rId = masterId.Attribute(OfficeNs + "id")?.Value;
            if (string.IsNullOrEmpty(rId)) continue;

            var masterPath = ResolveRelationship(doc, "ppt/presentation.xml", rId);
            if (string.IsNullOrEmpty(masterPath)) continue;

            var masterXml = doc.GetXml(masterPath);
            if (masterXml.Root != null)
            {
                presentation.MainMaster = masterXml;
            }
        }

        var themeRels = doc.GetRelationships("ppt/presentation.xml");
        foreach (var themeRel in themeRels)
        {
            if (themeRel.Contains("theme"))
            {
                var themeXml = doc.GetXml("ppt/" + themeRel);
                if (themeXml.Root != null)
                {
                    presentation.ThemeDocuments[themeRel] = themeXml;
                }
            }
        }

        foreach (var entry in doc.GetAllFiles())
        {
            if (entry.StartsWith("ppt/media/"))
            {
                var data = doc.GetFile(entry);
                presentation.MediaFiles[entry] = data;
            }
        }

        foreach (var slide in presentation.Slides)
        {
            if (slide.Xml?.Root != null)
            {
                ParseFontsFromSlide(slide.Xml, presentation);
            }
        }

        return presentation;
    }

    private void ParseFontsFromSlide(XDocument slideXml, PptxPresentation presentation)
    {
        if (slideXml.Root == null)
            return;
            
        var ns = slideXml.Root.GetDefaultNamespace();
        
        var textBodies = slideXml.Descendants(ns + "txBody");
        foreach (var txBody in textBodies)
        {
            // paragraphs are typically in the drawing namespace (<a:p>).  Use local
            // name comparison to find them.
            var paragraphs = txBody.Elements().Where(e => e.Name.LocalName == "p");
            foreach (var para in paragraphs)
            {
                // runs may also be in drawing namespace
                var runs = para.Elements().Where(e => e.Name.LocalName == "r");
                foreach (var run in runs)
                {
                    var rPr = run.Elements().FirstOrDefault(e => e.Name.LocalName == "rPr");
                    if (rPr != null)
                    {
                        var latin = rPr.Elements().FirstOrDefault(e => e.Name.LocalName == "latin");
                        if (latin != null)
                        {
                            var fontFace = latin.Attribute("typeface")?.Value;
                            if (!string.IsNullOrEmpty(fontFace) && !presentation.Fonts.ContainsKey(fontFace))
                            {
                                presentation.Fonts[fontFace] = presentation.Fonts.Count;
                            }
                        }
                    }
                }
            }
        }
    }

    private string ResolveRelationship(PptxDocument doc, string basePath, string rId)
    {
        var relDoc = doc.GetXml(GetRelPath(basePath));
        if (relDoc.Root == null)
            return "";

        var relNs = relDoc.Root.GetDefaultNamespace();
        var rel = relDoc.Root.Elements(relNs + "Relationship")
            .FirstOrDefault(r => r.Attribute("Id")?.Value == rId);

        var target = rel?.Attribute("Target")?.Value;
        if (string.IsNullOrEmpty(target))
            return "";

        if (target.StartsWith("/"))
            return target.TrimStart('/');

        var dir = Path.GetDirectoryName(basePath)?.Replace('\\', '/');
        if (string.IsNullOrEmpty(dir))
            return target;

        return dir + "/" + target;
    }

    private string GetRelPath(string basePath)
    {
        var dir = Path.GetDirectoryName(basePath)?.Replace('\\', '/') ?? "";
        var name = Path.GetFileName(basePath);
        return dir + "/_rels/" + name + ".rels";
    }
}
