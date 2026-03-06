using System.Text;
using System.Xml.Linq;
using Nedev.FileConverters.PptxToPpt.Pptx;

namespace Nedev.FileConverters.PptxToPpt.Ppt;

public sealed class PptDocumentBuilder
{
    private readonly PptDocument _document = new();
    private int _slideIdCounter = 0x1000;
    private int _shapeIdCounter = 0x1000;
    private bool _masterAdded = false;
    private readonly HashSet<int> _layoutIndexesAdded = new();

    public void AddSlide(PptxSlide slide)
    {
        var slideRecord = new SlideRecord
        {
            SlideId = _slideIdCounter++,
            Index = slide.Index
        };

        var shapeRecords = CreateShapesFromSlide(slide, null);
        slideRecord.Records.AddRange(shapeRecords);

        if (slide.NotesXml.TryGetValue(slide.Index, out var notesXml))
        {
            slideRecord.NotesData = CreateNotesData(notesXml);
        }

        _document.Slides.Add(slideRecord);
    }

    public void AddMaster(XDocument? master)
    {
        if (master == null || _masterAdded)
            return;

        var masterRecord = CreateMasterRecord(master);
        if (masterRecord != null)
        {
            _document.SlideMasters.Add(masterRecord);
            _masterAdded = true;
        }
    }

    public void AddLayouts(Dictionary<int, XDocument>? layouts)
    {
        if (layouts == null)
            return;

        foreach (var layout in layouts)
        {
            if (_layoutIndexesAdded.Contains(layout.Key))
                continue;

            var layoutRecord = CreateLayoutRecord(layout.Value);
            if (layoutRecord != null)
            {
                _document.SlideLayouts.Add(layout.Key, layoutRecord);
                _layoutIndexesAdded.Add(layout.Key);
            }
        }
    }

    private SlideRecord? CreateMasterRecord(XDocument masterXml)
    {
        if (masterXml.Root == null)
            return null;

        var record = new SlideRecord
        {
            SlideId = _slideIdCounter++,
            Index = 0
        };

        var ns = masterXml.Root.GetDefaultNamespace();
        var shapes = masterXml.Root.Descendants(ns + "sp").ToList();

        foreach (var shapeXml in shapes)
        {
            var shapeRecord = CreateShapeRecord(shapeXml);
            if (shapeRecord != null)
                record.Records.Add(shapeRecord);
        }

        return record;
    }

    private SlideRecord? CreateLayoutRecord(XDocument layoutXml)
    {
        if (layoutXml.Root == null)
            return null;

        var record = new SlideRecord
        {
            SlideId = _slideIdCounter++,
            Index = 0
        };

        var ns = layoutXml.Root.GetDefaultNamespace();
        var shapes = layoutXml.Root.Descendants(ns + "sp").ToList();

        foreach (var shapeXml in shapes)
        {
            var shapeRecord = CreateShapeRecord(shapeXml);
            if (shapeRecord != null)
                record.Records.Add(shapeRecord);
        }

        return record;
    }

    public void AddMedia(string name, byte[] data)
    {
        _document.MediaFiles[name] = data;
    }

    public ushort AddFont(string name, ushort charSet = 1, ushort family = 2)
    {
        if (!_document.Fonts.ContainsKey(name))
        {
            _document.Fonts[name] = new FontEntity
            {
                Name = name,
                CharSet = charSet,
                Family = family
            };
        }
        // index is based on enumeration order; convert keys to list to find position
        return (ushort)_document.Fonts.Keys.ToList().IndexOf(name);
    }

    public void WriteTo(Stream stream)
    {
        using var pptWriter = new PptWriter(stream);
        pptWriter.WriteDocument(_document);
    }

    private List<Record> CreateShapesFromSlide(PptxSlide slide, XDocument? master)
    {
        var records = new List<Record>();

        if (slide.Xml?.Root == null)
            return records;

        var ns = slide.Xml.Root.GetDefaultNamespace();

        var shapes = slide.Xml.Root.Descendants(ns + "sp").ToList();

        foreach (var shapeXml in shapes)
        {
            var shapeRecord = CreateShapeRecord(shapeXml);
            if (shapeRecord != null)
                records.Add(shapeRecord);
        }

        var groupShapes = slide.Xml.Root.Descendants(ns + "grpSp").ToList();
        foreach (var groupXml in groupShapes)
        {
            var groupRecord = CreateGroupShapeRecord(groupXml);
            if (groupRecord != null)
                records.Add(groupRecord);
        }
        // connectors (cxnSp) are treated similar to shapes
        var connectors = slide.Xml.Root.Descendants(ns + "cxnSp").ToList();
        foreach (var cxnXml in connectors)
        {
            var conn = CreateConnectorRecord(cxnXml);
            if (conn != null) records.Add(conn);
        }

        var pictures = slide.Xml.Root.Descendants(ns + "pic").ToList();
        foreach (var picXml in pictures)
        {
            var picRecord = CreatePictureRecord(picXml);
            if (picRecord != null)
                records.Add(picRecord);
        }

        return records;
    }

    private Record? CreatePictureRecord(XElement picXml)
    {
        var ns = picXml.GetDefaultNamespace();

        var nvPicPr = picXml.Element(ns + "nvPicPr");
        var name = nvPicPr?.Element(ns + "cNvPr")?.Attribute("name")?.Value ?? "";

        var blipFill = picXml.Element(ns + "blipFill");
        var blip = blipFill?.Element(ns + "blip");
        var embedAttr = blip?.Attribute("r:embed")?.Value;

        var picAtom = new Record
        {
            Type = RecordType.RT_PictureAtom,
            Version = 0x0FC3,
            Data = CreatePictureAtomData()
        };

        return new Record
        {
            Type = RecordType.RT_Picture,
            Version = 0x0FC2,
            Data = picAtom.ToArray()
        };
    }

    private byte[] CreatePictureAtomData()
    {
        var data = new byte[24];
        BitConverter.GetBytes((uint)0).CopyTo(data, 0);
        BitConverter.GetBytes((uint)_shapeIdCounter++).CopyTo(data, 4);
        BitConverter.GetBytes((uint)0).CopyTo(data, 8);
        BitConverter.GetBytes((uint)0x00060000).CopyTo(data, 12);
        BitConverter.GetBytes((uint)0x00010000).CopyTo(data, 16);
        BitConverter.GetBytes((uint)0x00000001).CopyTo(data, 20);
        return data;
    }

    private Record CreateShapeRecord(XElement shapeXml)
    {
        var ns = shapeXml.GetDefaultNamespace();

        var nvSpPr = shapeXml.Element(ns + "nvSpPr");
        var name = nvSpPr?.Element(ns + "cNvPr")?.Attribute("name")?.Value ?? "";

        var txBody = shapeXml.Element(ns + "txBody");

        var shapeAtom = new Record
        {
            Type = RecordType.RT_Shape,
            Version = 0x0FEC,
            Data = CreateShapeAtomData(shapeXml)
        };

        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        writer.Write(shapeAtom.ToArray());

        if (txBody != null)
        {
            var textRecords = CreateTextRecords(txBody);
            foreach (var tr in textRecords)
            {
                writer.Write(tr.ToArray());
            }
        }

        return new Record
        {
            Type = RecordType.RT_Container,
            Version = 0x0FFF,
            Data = ms.ToArray()
        };
    }

    private byte[] CreateShapeAtomData(XElement shapeXml)
    {
        // 32‑byte shape atom header.  Bytes 12/16 currently hold location data in
        // PPT records, byte 20 is rotation (1/60000ths of a degree), and bytes
        // 24/28 are usually used for extents.  We default the flags at offset 12
        // similar to how they were originally hard‑coded.
        var data = new byte[32];

        BitConverter.GetBytes((uint)0).CopyTo(data, 0);
        BitConverter.GetBytes((uint)_shapeIdCounter++).CopyTo(data, 4);
        BitConverter.GetBytes((uint)0).CopyTo(data, 8);
        BitConverter.GetBytes((uint)0x000A0000).CopyTo(data, 12);
        BitConverter.GetBytes((uint)0x00010000).CopyTo(data, 16);

        // shape properties may use DrawingML namespace prefixes, so traverse by LocalName
        var spPr = shapeXml.Elements().FirstOrDefault(e => e.Name.LocalName == "spPr");
        if (spPr != null)
        {
            var xfrm = spPr.Elements().FirstOrDefault(e => e.Name.LocalName == "xfrm");
            if (xfrm != null)
            {
                // parse offset
                var off = xfrm.Elements().FirstOrDefault(e => e.Name.LocalName == "off");
                if (off != null)
                {
                    if (int.TryParse(off.Attribute("x")?.Value, out var x))
                        BitConverter.GetBytes(x).CopyTo(data, 12);
                    if (int.TryParse(off.Attribute("y")?.Value, out var y))
                        BitConverter.GetBytes(y).CopyTo(data, 16);
                }

                // rotation attribute (same semantics as group)
                var rot = xfrm.Attribute("rot")?.Value;
                if (!string.IsNullOrEmpty(rot) && int.TryParse(rot, out int rotation))
                {
                    BitConverter.GetBytes(rotation).CopyTo(data, 20);
                }

                // extents (scale)
                var ext = xfrm.Elements().FirstOrDefault(e => e.Name.LocalName == "ext");
                if (ext != null)
                {
                    if (int.TryParse(ext.Attribute("cx")?.Value, out var cx))
                        BitConverter.GetBytes(cx).CopyTo(data, 24);
                    if (int.TryParse(ext.Attribute("cy")?.Value, out var cy))
                        BitConverter.GetBytes(cy).CopyTo(data, 28);
                }
            }
        }

        return data;
    }

    private Record CreateConnectorRecord(XElement cxnXml)
    {
        // treat connector as a simple shape for now; ID incremented etc.
        var data = new byte[24];
        BitConverter.GetBytes((uint)0).CopyTo(data, 0);
        BitConverter.GetBytes((uint)_shapeIdCounter++).CopyTo(data, 4);
        BitConverter.GetBytes((uint)0).CopyTo(data, 8);
        BitConverter.GetBytes((uint)0x00180000).CopyTo(data, 12);
        BitConverter.GetBytes((uint)0).CopyTo(data, 16);
        BitConverter.GetBytes((uint)0).CopyTo(data, 20);
        return new Record
        {
            Type = RecordType.RT_Shape,
            Version = 0x0FC8,
            Data = data
        };
    }

    private Record CreateGroupShapeRecord(XElement groupXml)
    {
        // create header for the group shape itself; we will encode translation,
        // rotation and optional scale into the first 32 bytes so that downstream
        // code can treat the data similarly to individual shapes.  Translation
        // uses offsets 12/16, rotation sits at 20 and extents (scale) are stored
        // at 24/28.  This required extending the header from 24 to 32 bytes.
        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        var header = new byte[32];
        BitConverter.GetBytes((uint)0).CopyTo(header, 0);
        BitConverter.GetBytes((uint)_shapeIdCounter++).CopyTo(header, 4);
        BitConverter.GetBytes((uint)0).CopyTo(header, 8);
        // default values (flags/offset) – will be overwritten if transform present
        BitConverter.GetBytes((uint)0x00180000).CopyTo(header, 12);
        BitConverter.GetBytes((uint)0).CopyTo(header, 16);
        BitConverter.GetBytes((uint)0).CopyTo(header, 20); // rotation default
        BitConverter.GetBytes((uint)0).CopyTo(header, 24); // ext cx default
        BitConverter.GetBytes((uint)0).CopyTo(header, 28); // ext cy default

        // parse transform if present.  We cannot rely on the default namespace
        // when navigating the drawingML-provided <a:grpSpPr>/<a:xfrm> tree, so use
        // LocalName comparisons similar to the text-parsing helpers earlier in the
        // file.
        var grpPr = groupXml.Elements().FirstOrDefault(e => e.Name.LocalName == "grpSpPr");
        var xfrm = grpPr?.Elements().FirstOrDefault(e => e.Name.LocalName == "xfrm");
        if (xfrm != null)
        {
            var off = xfrm.Elements().FirstOrDefault(e => e.Name.LocalName == "off");
            if (off != null)
            {
                if (int.TryParse(off.Attribute("x")?.Value, out var x))
                    BitConverter.GetBytes(x).CopyTo(header, 12);
                if (int.TryParse(off.Attribute("y")?.Value, out var y))
                    BitConverter.GetBytes(y).CopyTo(header, 16);
            }

            var rotAttr = xfrm.Attribute("rot")?.Value;
            if (!string.IsNullOrEmpty(rotAttr) && int.TryParse(rotAttr, out var rotation))
            {
                BitConverter.GetBytes(rotation).CopyTo(header, 20);
            }

            var ext = xfrm.Elements().FirstOrDefault(e => e.Name.LocalName == "ext");
            if (ext != null)
            {
                if (int.TryParse(ext.Attribute("cx")?.Value, out var cx))
                    BitConverter.GetBytes(cx).CopyTo(header, 24);
                if (int.TryParse(ext.Attribute("cy")?.Value, out var cy))
                    BitConverter.GetBytes(cy).CopyTo(header, 28);
            }
        }

        writer.Write(header);

        // recurse into contained elements.  We still need the default namespace of
        // the group itself for the Descendants calls below, so grab it here.
        var ns = groupXml.GetDefaultNamespace();
        foreach (var sp in groupXml.Descendants(ns + "sp"))
        {
            var rec = CreateShapeRecord(sp);
            if (rec != null) writer.Write(rec.ToArray());
        }
        foreach (var gp in groupXml.Descendants(ns + "grpSp"))
        {
            // avoid infinite recursion by ensuring it's a direct child
            if (gp != groupXml)
            {
                var rec = CreateGroupShapeRecord(gp);
                if (rec != null) writer.Write(rec.ToArray());
            }
        }
        foreach (var pic in groupXml.Descendants(ns + "pic"))
        {
            var rec = CreatePictureRecord(pic);
            if (rec != null) writer.Write(rec.ToArray());
        }
        foreach (var cxn in groupXml.Descendants(ns + "cxnSp"))
        {
            var rec = CreateConnectorRecord(cxn);
            if (rec != null) writer.Write(rec.ToArray());
        }

        return new Record
        {
            Type = RecordType.RT_GroupShape,
            Version = 0x0FC8,
            Data = ms.ToArray()
        };
    }

    private List<Record> CreateTextRecords(XElement txBody)
    {
        var records = new List<Record>();
        // paragraphs are in the drawing namespace (<a:p>) so default namespace
        // of txBody is not reliable.  Look for any child element whose local
        // name is "p".
        var paragraphs = txBody.Elements().Where(e => e.Name.LocalName == "p").ToList();

        // counters for auto-number by level
        var counters = new int[10];
        int prevLevel = 0;
        var defaultGlyphs = new[] { "•", "○", "▪", "–" };

        foreach (var para in paragraphs)
        {
            var pPr = para.Elements().FirstOrDefault(e => e.Name.LocalName == "pPr");
            int level = 0;
            string prefix = string.Empty;

            if (pPr != null)
            {
                var lvlAttr = pPr.Attribute("lvl")?.Value;
                if (!string.IsNullOrEmpty(lvlAttr) && int.TryParse(lvlAttr, out var lvl))
                    level = lvl;

                // handle automatic numbering
                var auto = pPr.Elements().FirstOrDefault(e => e.Name.LocalName == "buAutoNum");
                if (auto != null)
                {
                    // reset deeper counters when level decreases
                    if (level <= prevLevel)
                    {
                        for (int i = level + 1; i < counters.Length; i++)
                            counters[i] = 0;
                    }
                    counters[level]++;
                    var type = auto.Attribute("type")?.Value;
                    prefix = FormatAuto(type, counters[level]);
                }

                // explicit bullet character
                var buChar = pPr.Elements().FirstOrDefault(e => e.Name.LocalName == "buChar");
                if (buChar != null)
                {
                    prefix = buChar.Attribute("char")?.Value ?? string.Empty;
                }

                prevLevel = level;
            }

            // if still no prefix, supply default glyph by level
            if (string.IsNullOrEmpty(prefix))
            {
                int idx = Math.Min(level, defaultGlyphs.Length - 1);
                prefix = defaultGlyphs[idx];
            }

            records.Add(CreateParagraphRecord(para, prefix));
        }

        return records;
    }

    private Record CreateParagraphRecord(XElement para, string bulletPrefix = "")
    {
        // paragraph element may be in a namespace such as drawing;
        // ignore namespace when looking for runs and text nodes.
        string bulletChar = bulletPrefix;
        string alignment = string.Empty;
        int level = 0;
        var pPr = para.Elements().FirstOrDefault(e => e.Name.LocalName == "pPr");
        if (pPr != null)
        {
            // explicit bullet character wins over prefix
            var buChar = pPr.Elements().FirstOrDefault(e => e.Name.LocalName == "buChar");
            if (buChar != null)
            {
                bulletChar = buChar.Attribute("char")?.Value ?? string.Empty;
            }

            // paragraph alignment attribute
            alignment = pPr.Attribute("algn")?.Value ?? string.Empty;

            var lvlAttr = pPr.Attribute("lvl")?.Value;
            if (!string.IsNullOrEmpty(lvlAttr) && int.TryParse(lvlAttr, out var lvl))
                level = lvl;
        }

        var paraFormat = new Record
        {
            Type = RecordType.RT_TextParaFormatAtom,
            Version = 0x0FEA,
            Data = CreateParaFormatData(alignment, level)
        };

        // gather runs by local name "r" regardless of namespace
        var runs = para.Elements().Where(e => e.Name.LocalName == "r").ToList();

        var headerAtom = new Record
        {
            Type = RecordType.RT_TextHeaderAtom,
            Version = 0x03E3,
            Data = CreateTextHeaderData()
        };

        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        writer.Write(paraFormat.ToArray());
        writer.Write(headerAtom.ToArray());

        bool firstRun = true;
        foreach (var run in runs)
        {
            // get text content ignoring namespace
            var t = run.Elements().FirstOrDefault(e => e.Name.LocalName == "t");
            var runText = t != null ? t.Value : string.Empty;

            // prepend bullet character on first run if present
            if (firstRun && !string.IsNullOrEmpty(bulletChar))
            {
                runText = bulletChar + " " + runText;
            }
            firstRun = false;

            // determine formatting information from the run
            var (fontIndex, bold, italic, fontSize, colorRgb, underline) = GetRunFormat(run);

            // write a char format atom for this run
            var charFormat = new Record
            {
                Type = RecordType.RT_TextCharFormatAtom,
                Version = 0x03E4,
                Data = CreateTextCharFormatData(runText.Length, fontIndex, bold, italic, fontSize, colorRgb, underline)
            };
            writer.Write(charFormat.ToArray());

            // write the actual text for the run
            var bytesAtom = new Record
            {
                Type = RecordType.RT_TextBytesAtom,
                Version = 0x03E6,
                Data = CreateTextBytesData(runText)
            };
            writer.Write(bytesAtom.ToArray());
        }

        return new Record
        {
            Type = RecordType.RT_TextParagraph,
            Version = 0x0FC8,
            Data = ms.ToArray()
        };
    }

    private byte[] CreateParaFormatData(string alignment = "", int level = 0)
    {
        var data = new byte[12];
        // byte 0 encodes alignment: left=0, center=1, right=2, justify=3
        switch (alignment)
        {
            case "ctr": data[0] = 1; break;
            case "r": data[0] = 2; break;
            case "just": data[0] = 3; break;
            default: data[0] = 0; break; // left or unspecified
        }
        // byte 1 may hold indentation level (not used by PPT writer yet)
        data[1] = (byte)level;

        // remaining bytes currently hardcoded
        data[2] = 0x00;
        data[3] = 0x00;
        BitConverter.GetBytes((uint)0x00000001).CopyTo(data, 4);
        BitConverter.GetBytes((uint)0x00000000).CopyTo(data, 8);
        return data;
    }

    // helper for automatic list prefixes
    private string FormatAuto(string? type, int count)
    {
        if (string.IsNullOrEmpty(type))
            return count + ".";

        switch (type)
        {
            case "alphaLc":
                return ((char)('a' + (count - 1) % 26)) + ".";
            case "alphaUc":
                return ((char)('A' + (count - 1) % 26)) + ".";
            case "romanLc":
                return ToRoman(count).ToLowerInvariant() + ".";
            case "romanUc":
                return ToRoman(count).ToUpperInvariant() + ".";
            default:
                return count + ".";
        }
    }

    private string ToRoman(int number)
    {
        if (number < 1) return string.Empty;
        var numerals = new[]
        {
            Tuple.Create(1000, "M"), Tuple.Create(900, "CM"), Tuple.Create(500, "D"), Tuple.Create(400, "CD"),
            Tuple.Create(100, "C"), Tuple.Create(90, "XC"), Tuple.Create(50, "L"), Tuple.Create(40, "XL"),
            Tuple.Create(10, "X"), Tuple.Create(9, "IX"), Tuple.Create(5, "V"), Tuple.Create(4, "IV"), Tuple.Create(1, "I")
        };
        var result = new StringBuilder();
        foreach (var pair in numerals)
        {
            while (number >= pair.Item1)
            {
                result.Append(pair.Item2);
                number -= pair.Item1;
            }
        }
        return result.ToString();
    }

    private byte[] CreateTextHeaderData()
    {
        var data = new byte[8];
        data[0] = 0x00;
        data[1] = 0x00;
        data[2] = 0x00;
        data[3] = 0x00;
        BitConverter.GetBytes((uint)0x00000001).CopyTo(data, 4);
        return data;
    }

    /// <summary>
    /// Creates the binary payload for a TextCharFormatAtom.  We support
    /// a small subset of formatting: font index, size, bold, italic,
    /// underline, and an RGB color value.  Additional flags and fields
    /// are reserved for future expansion.
    /// </summary>
    private byte[] CreateTextCharFormatData(int runLength, ushort fontIndex, bool bold, bool italic, ushort fontSize = 0, uint colorRgb = 0, bool underline = false)
    {
        // Extended layout: [length:4][fontIdx:2][size:2][flags:2][reserved:2][color:4]
        var data = new byte[18];
        BitConverter.GetBytes((uint)runLength).CopyTo(data, 0);
        BitConverter.GetBytes(fontIndex).CopyTo(data, 4);
        BitConverter.GetBytes(fontSize).CopyTo(data, 6);
        ushort flags = 0;
        if (bold) flags |= 0x0001;
        if (italic) flags |= 0x0002;
        if (underline) flags |= 0x0004; // new flag for underline
        BitConverter.GetBytes(flags).CopyTo(data, 8);
        // bytes 10..11 reserved
        BitConverter.GetBytes(colorRgb).CopyTo(data, 12);
        return data;
    }

    /// <summary>
    /// Inspect a &lt;r&gt; run element and determine the corresponding
    /// font index (ensuring the font is registered) and style flags.
    /// </summary>
    private (ushort fontIndex, bool bold, bool italic, ushort size, uint colorRgb, bool underline) GetRunFormat(XElement run)
    {
        bool bold = false;
        bool italic = false;
        bool underline = false;
        ushort fontIndex = 0;
        ushort size = 0;
        uint colorRgb = 0;

        // ignore namespace when inspecting run properties
        var rPr = run.Elements().FirstOrDefault(e => e.Name.LocalName == "rPr");
        if (rPr != null)
        {
            var bAttr = rPr.Attribute("b")?.Value;
            if (!string.IsNullOrEmpty(bAttr) && (bAttr == "1" || bAttr.Equals("true", StringComparison.OrdinalIgnoreCase)))
                bold = true;
            var iAttr = rPr.Attribute("i")?.Value;
            if (!string.IsNullOrEmpty(iAttr) && (iAttr == "1" || iAttr.Equals("true", StringComparison.OrdinalIgnoreCase)))
                italic = true;

            // underline attribute (val may be "sng", "dbl", etc.).
            var uAttr = rPr.Attribute("u")?.Value;
            if (!string.IsNullOrEmpty(uAttr) && !uAttr.Equals("none", StringComparison.OrdinalIgnoreCase))
                underline = true;

            var szAttr = rPr.Attribute("sz")?.Value;
            if (!string.IsNullOrEmpty(szAttr) && ushort.TryParse(szAttr, out var szValue))
            {
                size = szValue;
            }

            var latin = rPr.Elements().FirstOrDefault(e => e.Name.LocalName == "latin");
            if (latin != null)
            {
                var face = latin.Attribute("typeface")?.Value;
                if (!string.IsNullOrEmpty(face))
                {
                    fontIndex = AddFont(face);
                }
            }

            // parse color if present (we only look for simple srgbClr values)
            var solidFill = rPr.Elements().FirstOrDefault(e => e.Name.LocalName == "solidFill");
            if (solidFill != null)
            {
                var srgb = solidFill.Descendants().FirstOrDefault(e => e.Name.LocalName == "srgbClr");
                var val = srgb?.Attribute("val")?.Value;
                if (!string.IsNullOrEmpty(val) && uint.TryParse(val, System.Globalization.NumberStyles.HexNumber, null, out var rgb))
                {
                    colorRgb = rgb;
                }
            }
        }

        return (fontIndex, bold, italic, size, colorRgb, underline);
    }

    private byte[] CreateTextBytesData(string text)
    {
        var textBytes = Encoding.UTF8.GetBytes(text);
        var data = new byte[4 + textBytes.Length];
        BitConverter.GetBytes((uint)textBytes.Length).CopyTo(data, 0);
        textBytes.CopyTo(data, 4);
        return data;
    }

    private byte[] CreateNotesData(XDocument notesXml)
    {
        if (notesXml.Root == null)
            return Array.Empty<byte>();

        var ns = notesXml.Root.GetDefaultNamespace();
        var txBody = notesXml.Root.Element(ns + "txBody");

        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        var notesAtom = new Record
        {
            Type = RecordType.RT_NotesAtom,
            Version = 0x03F5,
            Data = new byte[16]
        };
        writer.Write(notesAtom.ToArray());

        if (txBody != null)
        {
            var textRecords = CreateTextRecords(txBody);
            foreach (var tr in textRecords)
            {
                writer.Write(tr.ToArray());
            }
        }

        return new Record
        {
            Type = RecordType.RT_Notes,
            Version = 0x03F4,
            Data = ms.ToArray()
        }.ToArray();
    }
}
