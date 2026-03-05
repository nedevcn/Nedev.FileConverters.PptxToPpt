using System.Text;

namespace Nedev.PptxToPpt.Ppt;

public sealed class PptWriter : IDisposable
{
    private readonly Stream _stream;
    private readonly Cff.CffWriter _cff;
    private bool _disposed;

    public PptWriter(Stream stream)
    {
        _stream = stream;
        _cff = new Cff.CffWriter(stream);
    }

    public void WriteDocument(PptDocument document)
    {
        var documentDir = _cff.CreateDirectory("PowerPoint Document");
        documentDir.CLsid = new byte[16];
        documentDir.IsDirectory = true;
        
        var docStream = _cff.GetEntryStream(documentDir);
        var docWriter = new BinaryWriter(docStream, Encoding.UTF8, leaveOpen: true);

        var documentRecords = document.GetDocumentRecords();
        
        foreach (var record in documentRecords)
        {
            docWriter.Write(record);
        }

        docWriter.Flush();

        var currentUser = _cff.CreateDirectory("Current User");
        currentUser.CLsid = new byte[16];
        currentUser.IsDirectory = true;
        
        var userStream = _cff.GetEntryStream(currentUser);
        var userWriter = new BinaryWriter(userStream, Encoding.UTF8, leaveOpen: true);

        userWriter.Write(CreateCurrentUserAtom());
        userWriter.Write(CreateUserEditStg());
        userWriter.Flush();

        foreach (var media in document.MediaFiles)
        {
            var mediaDir = _cff.CreateDirectory("PowerPoint Document");
            
            var mediaEntry = _cff.CreateDirectory("PowerPoint Document/" + media.Key);
            mediaEntry.CLsid = new byte[16];
            mediaEntry.IsDirectory = false;
            
            var mediaStream = _cff.GetEntryStream(mediaEntry);
            mediaStream.Write(media.Value);
            mediaStream.Flush();
        }

        _cff.SetRootDirectory(documentDir);
        _cff.Write();
    }

    private byte[] CreateCurrentUserAtom()
    {
        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);
        
        writer.Write((ushort)RecordType.RT_CurrentUserAtom);
        writer.Write((ushort)0x0FF6);
        writer.Write((uint)20);
        writer.Write((uint)0x0006L);
        writer.Write(new byte[4]);
        writer.Write((uint)0xE05C0160);
        writer.Write((uint)0);
        writer.Write((uint)0);
        
        return ms.ToArray();
    }

    private byte[] CreateUserEditStg()
    {
        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);
        
        writer.Write((ushort)RecordType.RT_UserEditStg);
        writer.Write((ushort)0x0FF4);
        writer.Write((uint)20);
        writer.Write(new byte[20]);
        
        return ms.ToArray();
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _cff.Dispose();
    }
}

public sealed class PptDocument
{
    public List<SlideRecord> Slides { get; } = new();
    public Dictionary<int, SlideRecord> SlideLayouts { get; } = new();
    public List<SlideRecord> SlideMasters { get; } = new();
    public Dictionary<string, byte[]> MediaFiles { get; } = new();
    public Dictionary<string, FontEntity> Fonts { get; } = new();

    public byte[] GetDocumentRecords()
    {
        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        writer.Write(CreateDocumentContainer());

        return ms.ToArray();
    }

    private byte[] CreateDocumentContainer()
    {
        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        writer.Write((ushort)RecordType.RT_Document);
        writer.Write((ushort)0x0FF5);
        
        var children = new List<byte[]>();

        children.Add(CreateDocumentAtom());

        var slideList = CreateSlidePersistPtrAtom();
        if (slideList.Length > 0)
            children.Add(slideList);

        var masterList = CreateMasterPersistPtrAtom();
        if (masterList.Length > 0)
            children.Add(masterList);

        var fontList = CreateFontContainer();
        if (fontList.Length > 0)
            children.Add(fontList);

        children.Add(CreateDrawingContainer());

        long totalSize = 8;
        foreach (var child in children)
            totalSize += child.Length;

        writer.Write((uint)totalSize);
        
        writer.Write((ushort)RecordType.RT_Document);
        writer.Write((ushort)0x0FF5);

        foreach (var child in children)
        {
            writer.Write(child);
        }

        return ms.ToArray();
    }

    private byte[] CreateMasterPersistPtrAtom()
    {
        if (SlideMasters.Count == 0)
            return Array.Empty<byte>();

        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        writer.Write((ushort)RecordType.RT_MainMaster);
        writer.Write((ushort)0x03F0);

        var masterData = new List<byte[]>();
        uint offset = 0;

        foreach (var master in SlideMasters)
        {
            var data = master.ToArray();
            masterData.Add(CreatePersistPtrEntry(0x0E01 + masterData.Count, offset));
            offset += (uint)data.Length;
        }

        long totalSize = 4 + 2 + masterData.Sum(e => e.Length);
        
        writer.Write((uint)totalSize);
        writer.Write((ushort)masterData.Count);
        
        foreach (var entry in masterData)
        {
            writer.Write(entry);
        }

        return ms.ToArray();
    }

    private byte[] CreateFontContainer()
    {
        if (Fonts.Count == 0)
            return Array.Empty<byte>();

        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        writer.Write((ushort)RecordType.RT_FontEntity);
        writer.Write((ushort)0x0FBA);

        var fontData = new List<byte[]>();
        foreach (var font in Fonts.Values)
        {
            fontData.Add(font.ToArray());
        }

        long totalSize = 4 + fontData.Sum(f => f.Length);
        
        writer.Write((uint)totalSize);

        foreach (var data in fontData)
        {
            writer.Write(data);
        }

        return ms.ToArray();
    }

    private byte[] CreateDocumentAtom()
    {
        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);
        
        writer.Write((ushort)RecordType.RT_DocumentAtom);
        writer.Write((ushort)0x0FFE);
        writer.Write((uint)8);
        
        writer.Write((ushort)0);
        writer.Write((ushort)Slides.Count);
        
        return ms.ToArray();
    }

    private byte[] CreateSlidePersistPtrAtom()
    {
        if (Slides.Count == 0)
            return Array.Empty<byte>();

        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        writer.Write((ushort)RecordType.RT_SlidePersistAtom);
        writer.Write((ushort)0x1772);

        int offset = 0;
        var entries = new List<byte[]>();
        
        uint baseOffset = 0x08 + 0x08 + 4 + 2;
        
        foreach (var slide in Slides)
        {
            var slideData = slide.ToArray();
            var entry = CreatePersistPtrEntry(1 + offset, baseOffset + (uint)offset);
            entries.Add(entry);
            offset += slideData.Length;
        }

        int count = entries.Count;
        long totalSize = 4 + 2 + entries.Sum(e => e.Length);
        
        writer.Write((uint)totalSize);
        
        writer.Write((ushort)count);
        
        foreach (var entry in entries)
        {
            writer.Write(entry);
        }

        return ms.ToArray();
    }

    private byte[] CreatePersistPtrEntry(int persistId, uint offset)
    {
        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        int rawValue = (persistId << 20) | (int)offset;
        
        var bytes = new List<byte>();
        int temp = rawValue;
        do
        {
            bytes.Add((byte)((temp & 0x7F) | (bytes.Count > 0 ? 0x80 : 0)));
            temp >>= 7;
        } while (temp != 0);

        foreach (var b in bytes)
            writer.Write(b);

        return ms.ToArray();
    }

    private byte[] CreateDrawingContainer()
    {
        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        writer.Write((ushort)RecordType.RT_Drawing);
        writer.Write((ushort)0x0404);

        var drawings = new List<byte[]>();
        
        foreach (var slide in Slides)
        {
            var slideDrawing = slide.GetDrawingData();
            if (slideDrawing.Length > 0)
                drawings.Add(slideDrawing);
        }

        long totalSize = 4;
        foreach (var d in drawings)
            totalSize += d.Length;

        writer.Write((uint)totalSize);

        foreach (var d in drawings)
            writer.Write(d);

        return ms.ToArray();
    }
}

public class SlideRecord
{
    public int SlideId { get; set; }
    public int Index { get; set; }
    public List<Record> Records { get; } = new();
    public byte[]? NotesData { get; set; }

    public byte[] ToArray()
    {
        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        writer.Write(CreateSlideContainer());

        return ms.ToArray();
    }

    private byte[] CreateSlideContainer()
    {
        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        writer.Write((ushort)RecordType.RT_Slide);
        writer.Write((ushort)0x03EE);

        var children = new List<byte[]>();
        children.Add(CreateSlideAtom());
        
        var drawing = CreateDrawing();
        if (drawing != null)
            children.Add(drawing);

        if (NotesData != null && NotesData.Length > 0)
        {
            children.Add(NotesData);
        }

        long totalSize = 8;
        foreach (var child in children)
            totalSize += child.Length;

        writer.Write((uint)totalSize);
        
        writer.Write((ushort)RecordType.RT_Slide);
        writer.Write((ushort)0x03EE);

        foreach (var child in children)
            writer.Write(child);

        return ms.ToArray();
    }

    private byte[] CreateSlideAtom()
    {
        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);
        
        writer.Write((ushort)RecordType.RT_SlideAtom);
        writer.Write((ushort)0x0FEE);
        writer.Write((uint)20);
        
        writer.Write((uint)0);
        writer.Write((uint)SlideId);
        writer.Write((uint)0);
        writer.Write((uint)0);
        writer.Write((uint)1);
        
        return ms.ToArray();
    }

    private byte[]? CreateDrawing()
    {
        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        var shapes = new List<byte[]>();
        
        foreach (var record in Records)
        {
            if (record.Type == RecordType.RT_Shape)
            {
                shapes.Add(record.ToArray());
            }
            else if (record.Type == RecordType.RT_GroupShape)
            {
                shapes.Add(record.ToArray());
            }
        }

        if (shapes.Count == 0)
            return null;

        writer.Write((ushort)RecordType.RT_Drawing);
        writer.Write((ushort)0x0404);

        long totalSize = 4;
        foreach (var s in shapes)
            totalSize += s.Length;

        writer.Write((uint)totalSize);

        foreach (var s in shapes)
            writer.Write(s);

        return ms.ToArray();
    }

    public byte[] GetDrawingData()
    {
        return CreateDrawing() ?? Array.Empty<byte>();
    }
}

public class Record
{
    public RecordType Type { get; set; }
    public ushort Version { get; set; }
    public byte[]? Data { get; set; }

    public int TotalSize => 8 + (Data?.Length ?? 0);

    public byte[] ToArray()
    {
        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        writer.Write((ushort)Type);
        writer.Write(Version);

        if (Data != null && Data.Length > 0)
        {
            writer.Write(Data.Length);
            writer.Write(Data);
        }
        else
        {
            writer.Write(0);
        }

        return ms.ToArray();
    }

    public static Record CreateContainer(RecordType type, params Record[] children)
    {
        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        foreach (var child in children)
        {
            writer.Write(child.ToArray());
        }

        var data = ms.ToArray();

        return new Record
        {
            Type = type,
            Version = 0x0FFF,
            Data = data
        };
    }
}

public enum RecordType : ushort
{
    RT_Document = 0x03E8,
    RT_DocumentAtom = 0x03E9,
    RT_Slide = 0x03EE,
    RT_SlideAtom = 0x03EF,
    RT_SlideBase = 0x03F8,
    RT_SlideLayout = 0x03F2,
    RT_SlideMaster = 0x03F6,
    RT_SlideListWithText = 0x0FF0,
    RT_UserEditStg = 0x0FF4,
    RT_CurrentUserAtom = 0x0FF6,
    RT_PersistPtrFullAtom = 0x1772,
    RT_SlidePersistAtom = 0x0FF2,
    RT_Drawing = 0x0404,
    RT_Container = 0x0FFF,
    RT_TextHeaderAtom = 0x03E3,
    RT_TextCharFormatAtom = 0x03E4,
    RT_TextParaFormatAtom = 0x03E5,
    RT_TextBytesAtom = 0x03E6,
    RT_TextCFRunAtom = 0x03E7,
    RT_CString = 0x0FEE,
    RT_WideString = 0x0FEF,
    RT_Schedule = 0x0FC8,
    RT_MainMaster = 0x03F0,
    RT_Environment = 0x03FC,
    RT_SlideShowSlideInfoAtom = 0x0FC2,
    RT_SlideShowInfoAtom = 0x0FC0,
    RT_SlideRangeAtom = 0x0FC4,
    RT_Notes = 0x03F4,
    RT_NotesAtom = 0x03F5,
    RT_ExObjList = 0x0FC8,
    RT_ExObjListAtom = 0x0FC9,
    RT_ExOleObjAtom = 0x0FCE,
    RT_ExOleObj = 0x0FCD,
    RT_SrKinsoku = 0x0F08,
    RT_EndSrstAtom = 0x0F0A,
    RT_SST = 0x0FC2,
    RT_SSTAtom = 0x0FC3,
    RT_FontEntityAtom = 0x0FB9,
    RT_FontEntity = 0x0FBA,
    RT_CoreProperties = 0x0FC0,
    RT_DocumentContainer = 0x0FF5,
    RT_Shape = 0x0FEC,
    RT_GroupShape = 0x0FC8,
    RT_TextContainer = 0x0FC2,
    RT_TextParagraph = 0x0FC8,
    RT_Picture = 0x0FC2,
    RT_PictureAtom = 0x0FC3,
    RT_SlideBaseAtom = 0x03F9,
    RT_AnimationInfoAtom = 0x0444,
    RT_AnimationInfoContainer = 0x0445,
}

public class FontEntity
{
    public string Name { get; set; } = "";
    public ushort CharSet { get; set; }
    public ushort Family { get; set; }
    public byte[] ToArray()
    {
        var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        var nameBytes = Encoding.Unicode.GetBytes(Name + "\0");
        var paddedName = new byte[32];
        Array.Copy(nameBytes, paddedName, Math.Min(nameBytes.Length, 32));
        writer.Write(paddedName);

        writer.Write(CharSet);
        writer.Write(Family);
        writer.Write((byte)0);
        writer.Write((byte)0);
        writer.Write((ushort)0);

        return ms.ToArray();
    }
}
