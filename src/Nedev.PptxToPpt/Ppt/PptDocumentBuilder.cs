using System.Text;
using System.Xml.Linq;
using Nedev.PptxToPpt.Pptx;

namespace Nedev.PptxToPpt.Ppt;

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

    public void AddFont(string name, ushort charSet = 1, ushort family = 2)
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
        var data = new byte[32];

        BitConverter.GetBytes((uint)0).CopyTo(data, 0);
        BitConverter.GetBytes((uint)_shapeIdCounter++).CopyTo(data, 4);
        BitConverter.GetBytes((uint)0).CopyTo(data, 8);
        BitConverter.GetBytes((uint)0x000A0000).CopyTo(data, 12);
        BitConverter.GetBytes((uint)0x00010000).CopyTo(data, 16);

        var spPr = shapeXml.Element(shapeXml.GetDefaultNamespace() + "spPr");
        if (spPr != null)
        {
            var xfrm = spPr.Element(spPr.GetDefaultNamespace() + "xfrm");
            if (xfrm != null)
            {
                var rot = xfrm.Attribute("rot")?.Value;
                if (!string.IsNullOrEmpty(rot))
                {
                    if (int.TryParse(rot, out int rotation))
                    {
                        BitConverter.GetBytes(rotation).CopyTo(data, 20);
                    }
                }
            }
        }

        return data;
    }

    private Record CreateGroupShapeRecord(XElement groupXml)
    {
        var data = new byte[24];
        BitConverter.GetBytes((uint)0).CopyTo(data, 0);
        BitConverter.GetBytes((uint)_shapeIdCounter++).CopyTo(data, 4);
        BitConverter.GetBytes((uint)0).CopyTo(data, 8);
        BitConverter.GetBytes((uint)0x00180000).CopyTo(data, 12);
        BitConverter.GetBytes((uint)0).CopyTo(data, 16);
        BitConverter.GetBytes((uint)0).CopyTo(data, 20);

        return new Record
        {
            Type = RecordType.RT_GroupShape,
            Version = 0x0FC8,
            Data = data
        };
    }

    private List<Record> CreateTextRecords(XElement txBody)
    {
        var records = new List<Record>();
        var ns = txBody.GetDefaultNamespace();

        var paragraphs = txBody.Elements(ns + "p").ToList();

        foreach (var para in paragraphs)
        {
            records.Add(CreateParagraphRecord(para));
        }

        return records;
    }

    private Record CreateParagraphRecord(XElement para)
    {
        var ns = para.GetDefaultNamespace();

        var paraFormat = new Record
        {
            Type = RecordType.RT_TextParaFormatAtom,
            Version = 0x0FEA,
            Data = CreateParaFormatData()
        };

        var runs = para.Elements(ns + "r").ToList();
        var textContent = new StringBuilder();
        
        foreach (var run in runs)
        {
            var t = run.Element(ns + "t");
            if (t != null)
            {
                textContent.Append(t.Value);
            }
        }

        var text = textContent.ToString();

        var headerAtom = new Record
        {
            Type = RecordType.RT_TextHeaderAtom,
            Version = 0x03E3,
            Data = CreateTextHeaderData()
        };

        var bytesAtom = new Record
        {
            Type = RecordType.RT_TextBytesAtom,
            Version = 0x03E6,
            Data = CreateTextBytesData(text)
        };

        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);
        
        writer.Write(paraFormat.ToArray());
        writer.Write(headerAtom.ToArray());
        writer.Write(bytesAtom.ToArray());

        return new Record
        {
            Type = RecordType.RT_TextParagraph,
            Version = 0x0FC8,
            Data = ms.ToArray()
        };
    }

    private byte[] CreateParaFormatData()
    {
        var data = new byte[12];
        data[0] = 0x00;
        data[1] = 0x00;
        data[2] = 0x00;
        data[3] = 0x00;
        BitConverter.GetBytes((uint)0x00000001).CopyTo(data, 4);
        BitConverter.GetBytes((uint)0x00000000).CopyTo(data, 8);
        return data;
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
