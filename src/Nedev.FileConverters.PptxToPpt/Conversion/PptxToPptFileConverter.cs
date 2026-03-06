using System;
using System.IO;
using Nedev.FileConverters.Core;

namespace Nedev.FileConverters.PptxToPpt.Conversion
{
    /// <summary>
    /// Adapter that exposes the existing Converter logic through the
    /// <see cref="IFileConverter"/> interface from the core NuGet package.
    /// This allows the static Converter/registry in the core package to be used
    /// for PPTX-&gt;PPT conversions.
    /// </summary>
    [FileConverter("pptx", "ppt")]
    public sealed class PptxToPptFileConverter : IFileConverter
    {
        public Stream Convert(Stream input)
        {
            // write to temporary files and delegate to existing Converter
            var inTemp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".pptx");
            var outTemp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".ppt");

            try
            {
                using (var f = File.Create(inTemp))
                {
                    input.CopyTo(f);
                }

                var conv = new Converter();
                conv.Convert(inTemp, outTemp, overwrite: true);

                var ms = new MemoryStream();
                using (var outFs = File.OpenRead(outTemp))
                {
                    outFs.CopyTo(ms);
                }
                ms.Position = 0;
                return ms;
            }
            finally
            {
                try { File.Delete(inTemp); } catch { }
                try { File.Delete(outTemp); } catch { }
            }
        }
    }
}