using System.Text;

namespace Nedev.FileConverters.PptxToPpt.Cff;

public sealed class CffWriter : IDisposable
{
    private const int SectorSize = 512;
    private const int SectorSizePow2 = 9;
    private const int MiniSectorSize = 64;
    private const int MiniSectorSizePow2 = 6;
    private const uint Freesect = 0xFFFFFFFE;
    private const uint Endosect = 0xFFFFFFFD;
    private const uint FATSect = 0xFFFFFFFC;

    private readonly Stream _stream;
    private readonly BinaryWriter _writer;
    private readonly List<CffDirectoryEntry> _directories = new();
    private readonly Dictionary<string, int> _nameToIndex = new();
    private int _rootDirectoryIndex = -1;
    private int _fatSectorsCount = 0;
    private readonly List<byte[]> _fatSectors = new();
    private readonly List<byte[]> _dataSectors = new();
    private bool _disposed;

    public CffWriter(Stream stream)
    {
        _stream = stream;
        _writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true);
    }

    public CffDirectoryEntry CreateDirectory(string name)
    {
        var entry = new CffDirectoryEntry
        {
            Name = name,
            Index = _directories.Count
        };
        _directories.Add(entry);
        _nameToIndex[name] = entry.Index;
        return entry;
    }

    public CffDirectoryEntry GetDirectory(string name)
    {
        if (_nameToIndex.TryGetValue(name, out var index))
            return _directories[index];
        return CreateDirectory(name);
    }

    public void SetRootDirectory(CffDirectoryEntry root)
    {
        _rootDirectoryIndex = root.Index;
    }

    public Stream GetEntryStream(CffDirectoryEntry entry)
    {
        return new CffEntryStream(this, entry);
    }

    public void Write()
    {
        BuildFatChain();
        WriteHeader();
        WriteFatSectors();
        WriteDirectorySectors();
        WriteDataSectors();
    }

    private void BuildFatChain()
    {
        var totalDataSectors = 0;
        foreach (var dir in _directories)
        {
            if (dir.Data != null && dir.Data.Length > 0)
            {
                dir.SectorCount = (dir.Data.Length + SectorSize - 1) / SectorSize;
            }
            else
            {
                dir.SectorCount = 0;
            }
            totalDataSectors += dir.SectorCount;
        }

        int fatSectorsNeeded = (totalDataSectors * 4 + SectorSize - 1) / SectorSize;
        if (fatSectorsNeeded < 1) fatSectorsNeeded = 1;
        if (fatSectorsNeeded > 109) fatSectorsNeeded = 109;

        _fatSectorsCount = fatSectorsNeeded;

        for (int i = 0; i < fatSectorsNeeded; i++)
        {
            var fatSector = new byte[SectorSize];
            _fatSectors.Add(fatSector);
        }

        int nextFreeSector = 1 + _fatSectorsCount + 1;
        int dirSectorCount = (_directories.Count * 128 + SectorSize - 1) / SectorSize;
        
        for (int i = 0; i < dirSectorCount; i++)
        {
            _dataSectors.Add(new byte[SectorSize]);
        }
        nextFreeSector += dirSectorCount;

        foreach (var dir in _directories)
        {
            if (dir.SectorCount > 0)
            {
                dir.StartSector = (uint)nextFreeSector;
                for (int i = 0; i < dir.SectorCount - 1; i++)
                {
                    int sectorIndex = nextFreeSector + i;
                    SetFatEntry(sectorIndex, (uint)(sectorIndex + 1));
                }
                SetFatEntry(nextFreeSector + dir.SectorCount - 1, Endosect);
                nextFreeSector += dir.SectorCount;
            }
        }

        for (int i = 0; i < _fatSectorsCount; i++)
        {
            int fatSectorIndex = 1 + i;
            if (i < _fatSectorsCount - 1)
                SetFatEntry(fatSectorIndex, (uint)(fatSectorIndex + 1));
            else
                SetFatEntry(fatSectorIndex, Endosect);
        }

        int dirStartSector = 1 + _fatSectorsCount;
        for (int i = 0; i < dirSectorCount - 1; i++)
        {
            SetFatEntry(dirStartSector + i, (uint)(dirStartSector + i + 1));
        }
        if (dirSectorCount > 0)
            SetFatEntry(dirStartSector + dirSectorCount - 1, Endosect);
    }

    private void SetFatEntry(int sectorIndex, uint value)
    {
        int fatSectorIndex = sectorIndex / (SectorSize / 4);
        int offsetInFat = (sectorIndex % (SectorSize / 4)) * 4;

        if (fatSectorIndex < _fatSectors.Count)
        {
            BitConverter.GetBytes(value).CopyTo(_fatSectors[fatSectorIndex], offsetInFat);
        }
    }

    private void WriteHeader()
    {
        var header = new byte[512];

        header[0] = 0xD0;
        header[1] = 0xCF;
        header[2] = 0x11;
        header[3] = 0xE0;
        header[4] = 0xA1;
        header[5] = 0xB1;
        header[6] = 0x1A;
        header[7] = 0xE1;

        for (int i = 8; i < 16; i++)
            header[i] = 0x00;

        for (int i = 16; i < 24; i++)
            header[i] = 0xFE;

        header[24] = 0x00;
        header[25] = 0x00;
        header[26] = 0x00;
        header[27] = 0x00;
        header[28] = 0x00;
        header[29] = 0x00;
        header[30] = 0x00;
        header[31] = 0x00;

        header[32] = 0x00;
        header[33] = 0x00;
        header[34] = 0x00;
        header[35] = 0x00;
        header[36] = 0x00;
        header[37] = 0x00;
        header[38] = 0x00;
        header[39] = 0x00;

        BitConverter.GetBytes((ushort)0x003E).CopyTo(header, 0x18);
        BitConverter.GetBytes((ushort)0x0003).CopyTo(header, 0x1A);
        BitConverter.GetBytes((ushort)0xFFFE).CopyTo(header, 0x1C);
        BitConverter.GetBytes((ushort)0x0009).CopyTo(header, 0x1E);

        BitConverter.GetBytes((ushort)0x0000).CopyTo(header, 0x20);
        BitConverter.GetBytes((ushort)0x0000).CopyTo(header, 0x22);

        BitConverter.GetBytes((uint)0x00000000).CopyTo(header, 0x24);
        BitConverter.GetBytes((uint)0x00000000).CopyTo(header, 0x28);
        BitConverter.GetBytes((uint)0x00000000).CopyTo(header, 0x2C);

        BitConverter.GetBytes((uint)0x00100000).CopyTo(header, 0x30);
        BitConverter.GetBytes((uint)0x00000000).CopyTo(header, 0x34);
        BitConverter.GetBytes((uint)0x00000000).CopyTo(header, 0x38);

        BitConverter.GetBytes((uint)_fatSectorsCount).CopyTo(header, 0x3C);

        BitConverter.GetBytes((uint)0x00000000).CopyTo(header, 0x40);
        BitConverter.GetBytes((uint)0xFFFFFFFF).CopyTo(header, 0x44);

        BitConverter.GetBytes((uint)(1 + _fatSectorsCount)).CopyTo(header, 0x48);

        BitConverter.GetBytes((uint)0x00000000).CopyTo(header, 0x4C);
        BitConverter.GetBytes((uint)0xFFFFFFFF).CopyTo(header, 0x50);

        int dirSectorCount = (_directories.Count * 128 + SectorSize - 1) / SectorSize;
        BitConverter.GetBytes((uint)dirSectorCount).CopyTo(header, 0x54);

        BitConverter.GetBytes((uint)0x00000000).CopyTo(header, 0x58);
        BitConverter.GetBytes((uint)0x00000000).CopyTo(header, 0x5C);

        BitConverter.GetBytes((uint)0xFFFFFFFE).CopyTo(header, 0x60);
        BitConverter.GetBytes((uint)0xFFFFFFFE).CopyTo(header, 0x64);

        BitConverter.GetBytes((uint)0x00000000).CopyTo(header, 0x68);

        _writer.Write(header);
    }

    private void WriteFatSectors()
    {
        foreach (var fatSector in _fatSectors)
        {
            _writer.Write(fatSector);
        }
    }

    private void WriteDirectorySectors()
    {
        int dirSectorCount = (_directories.Count * 128 + SectorSize - 1) / SectorSize;
        
        for (int i = 0; i < dirSectorCount; i++)
        {
            var sectorData = new byte[128];
            int entriesInThisSector = Math.Min(_directories.Count - i * 4, 4);
            
            for (int j = 0; j < entriesInThisSector; j++)
            {
                var dir = _directories[i * 4 + j];
                var entryData = WriteDirectoryEntry(dir);
                Array.Copy(entryData, 0, sectorData, j * 128, 128);
            }
            
            _writer.Write(sectorData);
            _writer.Write(new byte[SectorSize - 128]);
        }
    }

    private byte[] WriteDirectoryEntry(CffDirectoryEntry dir)
    {
        var data = new byte[128];

        var nameBytes = Encoding.Unicode.GetBytes(dir.Name);
        Array.Copy(nameBytes, 0, data, 0, Math.Min(nameBytes.Length, 64));
        
        data[64] = (byte)Math.Min(dir.Name.Length, 32);
        data[65] = 0x00;
        
        if (dir.IsDirectory)
        {
            data[66] = 0x01;
        }
        else
        {
            data[66] = 0x02;
        }
        data[67] = 0x00;

        if (dir.LeftSibling >= 0)
            BitConverter.GetBytes((uint)dir.LeftSibling).CopyTo(data, 68);
        else
            BitConverter.GetBytes(Freesect).CopyTo(data, 68);

        if (dir.RightSibling >= 0)
            BitConverter.GetBytes((uint)dir.RightSibling).CopyTo(data, 72);
        else
            BitConverter.GetBytes(Freesect).CopyTo(data, 72);

        if (dir.Child >= 0)
            BitConverter.GetBytes((uint)dir.Child).CopyTo(data, 76);
        else
            BitConverter.GetBytes(Freesect).CopyTo(data, 76);

        data[80] = 0x00;
        for (int i = 81; i < 92; i++)
            data[i] = 0x00;

        BitConverter.GetBytes(dir.StateBits).CopyTo(data, 92);
        BitConverter.GetBytes(dir.CreationTime).CopyTo(data, 96);
        BitConverter.GetBytes(dir.ModifyTime).CopyTo(data, 100);

        BitConverter.GetBytes(dir.StartSector).CopyTo(data, 108);
        BitConverter.GetBytes(dir.Size).CopyTo(data, 112);

        BitConverter.GetBytes(dir.Size).CopyTo(data, 116);

        return data;
    }

    private void WriteDataSectors()
    {
        foreach (var dir in _directories)
        {
            if (dir.Data != null && dir.Data.Length > 0)
            {
                int offset = 0;
                for (int i = 0; i < dir.SectorCount; i++)
                {
                    var sector = new byte[SectorSize];
                    int copyLength = Math.Min(SectorSize, dir.Data.Length - offset);
                    if (copyLength > 0)
                    {
                        Array.Copy(dir.Data, offset, sector, 0, copyLength);
                    }
                    _writer.Write(sector);
                    offset += copyLength;
                }
            }
            else if (dir.SectorCount > 0)
            {
                for (int i = 0; i < dir.SectorCount; i++)
                {
                    _writer.Write(new byte[SectorSize]);
                }
            }
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _writer.Dispose();
    }
}

public class CffDirectoryEntry
{
    public string Name { get; set; } = "";
    public int Index { get; set; }
    public int LeftSibling { get; set; } = -1;
    public int RightSibling { get; set; } = -1;
    public int Child { get; set; } = -1;
    public byte[] CLsid { get; set; } = new byte[16];
    public uint StateBits { get; set; }
    public uint CreationTime { get; set; }
    public uint ModifyTime { get; set; }
    public uint StartSector { get; set; }
    public int Size { get; set; }
    public byte[]? Data { get; set; }
    public int SectorCount { get; set; }
    public bool IsDirectory { get; set; }
}

internal sealed class CffEntryStream : Stream
{
    private readonly CffWriter _writer;
    private readonly CffDirectoryEntry _entry;
    private readonly MemoryStream _memory;
    private bool _disposed;

    public CffEntryStream(CffWriter writer, CffDirectoryEntry entry)
    {
        _writer = writer;
        _entry = entry;
        _memory = new MemoryStream();
    }

    public override bool CanRead => false;
    public override bool CanSeek => _memory.CanSeek;
    public override bool CanWrite => true;
    public override long Length => _memory.Length;
    public override long Position { get => _memory.Position; set => _memory.Position = value; }

    public override void Write(byte[] buffer, int offset, int count)
    {
        _memory.Write(buffer, offset, count);
    }

    public override void Flush()
    {
        _entry.Data = _memory.ToArray();
        _entry.Size = _entry.Data.Length;
    }

    public override long Seek(long offset, SeekOrigin origin)
    {
        return _memory.Seek(offset, origin);
    }

    public override void SetLength(long value)
    {
        _memory.SetLength(value);
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        return _memory.Read(buffer, offset, count);
    }

    protected override void Dispose(bool disposing)
    {
        if (_disposed) return;
        if (disposing)
        {
            Flush();
            _memory.Dispose();
        }
        _disposed = true;
        base.Dispose(disposing);
    }
}
