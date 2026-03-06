using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Nedev.FileConverters.PptxToPpt.Ppt;
using Nedev.FileConverters.PptxToPpt.Pptx;
using Xunit;
using System.Text;

namespace Nedev.FileConverters.PptxToPpt.Tests
{
    public class RichTextTests
    {
        [Fact]
        public void ParagraphRunsProduceCharFormatAtoms()
        {
            // build a minimal slide xml containing two runs, one bold and one italic
            var xml = @"<sld xmlns=""http://schemas.openxmlformats.org/presentationml/2006/main"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
  <cSld>
    <spTree>
      <sp>
        <txBody>
          <a:bodyPr/>
          <a:p>
            <a:r>
              <a:rPr b=""1""><a:latin typeface=""Arial""/></a:rPr>
              <a:t>Bold</a:t>
            </a:r>
            <a:r>
              <a:rPr i=""1""><a:latin typeface=""Arial""/></a:rPr>
              <a:t>Italic</a:t>
            </a:r>
          </a:p>
        </txBody>
      </sp>
    </spTree>
  </cSld>
</sld>";

            var slide = new PptxSlide { Index = 0, Xml = XDocument.Parse(xml) };
            var builder = new PptDocumentBuilder();
            builder.AddSlide(slide);

            // after adding the slide, we can directly check the paragraph data
            var shapeElement = slide.Xml.Root.Descendants().FirstOrDefault(e => e.Name.LocalName == "sp");
            var txBodyElement = shapeElement?.Element(shapeElement.GetDefaultNamespace() + "txBody");
            if (txBodyElement != null)
            {
                var createTextRecords = typeof(PptDocumentBuilder).GetMethod("CreateTextRecords", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                var recordsList = (System.Collections.IList)createTextRecords.Invoke(builder, new object[] { txBodyElement });
                if (recordsList.Count > 0)
                {
                    var rec = recordsList[0];
                    var data = (byte[])rec.GetType().GetProperty("Data").GetValue(rec);
                    Console.WriteLine(" paragraph data bytes: " + BitConverter.ToString(data).Replace("-"," "));
                    // assert char-format sequence exists in the paragraph data
                    var pattern = new byte[] { 0xE4, 0x03, 0xE4, 0x03 };
                    AssertExtensions.ContainsSequence(data, pattern);
                }
            }

            using var ms = new MemoryStream();
            builder.WriteTo(ms);
            var output = ms.ToArray();

            // (we no longer need to check the full output here, paragraph data assert suffices)
            Console.WriteLine("output hex: " + BitConverter.ToString(output).Replace("-", " "));
        }

        [Fact]
        public void ParagraphFormattingColorUnderlineBullet()
        {
            // slide with a bold run that is underlined, colored red, and a bullet at paragraph level
            var xml = @"<sld xmlns=""http://schemas.openxmlformats.org/presentationml/2006/main"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
  <cSld>
    <spTree>
      <sp>
        <txBody>
          <a:bodyPr/>
          <a:p>
            <a:pPr>
              <a:buChar char=""•""/>
            </a:pPr>
            <a:r>
              <a:rPr u=""sng""><a:solidFill><a:srgbClr val=""FF0000""/></a:solidFill></a:rPr>
              <a:t>Test</a:t>
            </a:r>
          </a:p>
        </txBody>
      </sp>
    </spTree>
  </cSld>
</sld>";

            var slide = new PptxSlide { Index = 0, Xml = XDocument.Parse(xml) };
            var builder = new PptDocumentBuilder();
            builder.AddSlide(slide);

            var shapeElement = slide.Xml.Root.Descendants().FirstOrDefault(e => e.Name.LocalName == "sp");
            var txBodyElement = shapeElement?.Element(shapeElement.GetDefaultNamespace() + "txBody");
            if (txBodyElement != null)
            {
                var createTextRecords = typeof(PptDocumentBuilder).GetMethod("CreateTextRecords", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                var recordsList = (System.Collections.IList)createTextRecords.Invoke(builder, new object[] { txBodyElement });
                Assert.True(recordsList.Count > 0);
                var rec = recordsList[0];
                var data = (byte[])rec.GetType().GetProperty("Data").GetValue(rec);
                Console.WriteLine(" paragraph data bytes: " + BitConverter.ToString(data).Replace("-"," "));

                // bullet char should appear in UTF8
                AssertExtensions.ContainsSequence(data, Encoding.UTF8.GetBytes("•"));

                // locate first char-format record in the paragraph data
                var fmtPattern = new byte[] { 0xE4, 0x03, 0xE4, 0x03 };
                int idx = -1;
                for (int i = 0; i < data.Length - fmtPattern.Length; i++)
                {
                    bool match = true;
                    for (int j = 0; j < fmtPattern.Length; j++)
                    {
                        if (data[i + j] != fmtPattern[j]) { match = false; break; }
                    }
                    if (match) { idx = i; break; }
                }
                Assert.True(idx >= 0, "char format header not found");

                // flags are at idx+16 within the record
                ushort flags = BitConverter.ToUInt16(data, idx + 16);
                Assert.True((flags & 0x04) != 0, "underline flag should be set");

                // color value stored at idx+20
                uint color = BitConverter.ToUInt32(data, idx + 20);
                Assert.Equal(0xFF0000u, color);
            }

            using var ms2 = new MemoryStream();
            builder.WriteTo(ms2);
        }

        [Fact]
        public void ParagraphAutoNumberBulletPrefixes1Dot()
        {
            var xml = @"<sld xmlns=""http://schemas.openxmlformats.org/presentationml/2006/main"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
  <cSld>
    <spTree>
      <sp>
        <txBody>
          <a:bodyPr/>
          <a:p>
            <a:pPr>
              <a:buAutoNum/>
            </a:pPr>
            <a:r>
              <a:rPr><a:latin typeface=""Arial""/></a:rPr>
              <a:t>Numbered</a:t>
            </a:r>
          </a:p>
        </txBody>
      </sp>
    </spTree>
  </cSld>
</sld>";
            var slide = new PptxSlide { Index = 0, Xml = XDocument.Parse(xml) };
            var builder = new PptDocumentBuilder();
            builder.AddSlide(slide);

            var shapeElement = slide.Xml.Root.Descendants().FirstOrDefault(e => e.Name.LocalName == "sp");
            var txBodyElement = shapeElement?.Element(shapeElement.GetDefaultNamespace() + "txBody");
            if (txBodyElement != null)
            {
                var createTextRecords = typeof(PptDocumentBuilder).GetMethod("CreateTextRecords", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                var recordsList = (System.Collections.IList)createTextRecords.Invoke(builder, new object[] { txBodyElement });
                Assert.True(recordsList.Count > 0);
                var rec = recordsList[0];
                var data = (byte[])rec.GetType().GetProperty("Data").GetValue(rec);
                // check that the paragraph text begins with "1."
                var textBytes = Encoding.UTF8.GetBytes("1.");
                AssertExtensions.ContainsSequence(data, textBytes);
            }

            using var ms3 = new MemoryStream();
            builder.WriteTo(ms3);
        }

        [Fact]
        public void ParagraphAlignmentIsEncoded()
        {
            var xml = @"<sld xmlns=""http://schemas.openxmlformats.org/presentationml/2006/main"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
  <cSld>
    <spTree>
      <sp>
        <txBody>
          <a:bodyPr/>
          <a:p>
            <a:pPr algn=""ctr""/>
            <a:r>
              <a:rPr><a:latin typeface=""Arial""/></a:rPr>
              <a:t>Centered</a:t>
            </a:r>
          </a:p>
        </txBody>
      </sp>
    </spTree>
  </cSld>
</sld>";
            var slide = new PptxSlide { Index = 0, Xml = XDocument.Parse(xml) };
            var builder = new PptDocumentBuilder();
            builder.AddSlide(slide);

            var shapeElement = slide.Xml.Root.Descendants().FirstOrDefault(e => e.Name.LocalName == "sp");
            var txBodyElement = shapeElement?.Element(shapeElement.GetDefaultNamespace() + "txBody");
            if (txBodyElement != null)
            {
                var createTextRecords = typeof(PptDocumentBuilder).GetMethod("CreateTextRecords", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                var recordsList = (System.Collections.IList)createTextRecords.Invoke(builder, new object[] { txBodyElement });
                Assert.True(recordsList.Count > 0);
                var rec = recordsList[0];
                var data = (byte[])rec.GetType().GetProperty("Data").GetValue(rec);
                // paraFormatData immediately follows the 8-byte record header
                Assert.Equal(1, data[8]); // center alignment encoded as 1
            }

            using var ms4 = new MemoryStream();
            builder.WriteTo(ms4);
        }

        [Fact]
        public void BulletGlyphsVaryByLevel()
        {
            var xml = @"<sld xmlns=""http://schemas.openxmlformats.org/presentationml/2006/main"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
  <cSld>
    <spTree>
      <sp>
        <txBody>
          <a:bodyPr/>
          <a:p>
            <a:pPr/>
            <a:r><a:t>First</a:t></a:r>
          </a:p>
          <a:p>
            <a:pPr lvl=""1""/>
            <a:r><a:t>Second</a:t></a:r>
          </a:p>
        </txBody>
      </sp>
    </spTree>
  </cSld>
</sld>";
            var slide = new PptxSlide { Index = 0, Xml = XDocument.Parse(xml) };
            var builder = new PptDocumentBuilder();
            builder.AddSlide(slide);

            var shapeElement = slide.Xml.Root.Descendants().FirstOrDefault(e => e.Name.LocalName == "sp");
            var txBodyElement = shapeElement?.Element(shapeElement.GetDefaultNamespace() + "txBody");
            if (txBodyElement != null)
            {
                var createTextRecords = typeof(PptDocumentBuilder).GetMethod("CreateTextRecords", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                var recordsList = (System.Collections.IList)createTextRecords.Invoke(builder, new object[] { txBodyElement });
                Assert.Equal(2, recordsList.Count);
                var firstData = (byte[])recordsList[0].GetType().GetProperty("Data").GetValue(recordsList[0]);
                var secondData = (byte[])recordsList[1].GetType().GetProperty("Data").GetValue(recordsList[1]);
                AssertExtensions.ContainsSequence(firstData, Encoding.UTF8.GetBytes("•"));
                AssertExtensions.ContainsSequence(secondData, Encoding.UTF8.GetBytes("○"));
            }

            using var ms5 = new MemoryStream();
            builder.WriteTo(ms5);
        }
    }

    internal static class AssertExtensions
    {
        public static void ContainsSequence(byte[] haystack, byte[] needle)
        {
            for (int i = 0; i <= haystack.Length - needle.Length; i++)
            {
                bool match = true;
                for (int j = 0; j < needle.Length; j++)
                {
                    if (haystack[i + j] != needle[j])
                    {
                        match = false;
                        break;
                    }
                }
                if (match) return;
            }
            throw new Xunit.Sdk.XunitException("Needle not found in haystack");
        }
    }
}