using SkiaSharp;
using ZXing;
using ZXing.Common;

namespace ConsoleApp
{
    public static class Utils
    {
        public static void GenerateBarcodeToStream(string content, BarcodeFormat format, Stream outputStream, int width = 300, int height = 150)
        {
            var barcodeWriter = new BarcodeWriter<SKBitmap>
            {
                Format = format,
                Options = new EncodingOptions
                {
                    Width = width,
                    Height = height,
                    Margin = 1,
                    PureBarcode = true
                }
            };

            using SKBitmap barcodeBitmap = barcodeWriter.Write(content);
            using SKImage image = SKImage.FromBitmap(barcodeBitmap);
            using SKData data = image.Encode(SKEncodedImageFormat.Png, 100);
            data.SaveTo(outputStream);
        }

        public static MemoryStream GetBarcode(string content, BarcodeFormat format, int width = 300, int height = 150)
        {
            var ms = new MemoryStream();
            GenerateBarcodeToStream(content, format, ms, width, height);
            return ms;
        }
    }
}
