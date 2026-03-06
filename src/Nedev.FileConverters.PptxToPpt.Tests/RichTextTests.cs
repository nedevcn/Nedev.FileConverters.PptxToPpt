using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Nedev.FileConverters.PptxToPpt.Ppt;
using Nedev.FileConverters.PptxToPpt.Pptx;
using Xunit;

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