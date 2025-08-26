using SkiaSharp;
using ZXing;
using ZXing.Common;
using ZXing.Rendering;
using ZXing.SkiaSharp.Rendering;

namespace ConsoleApp
{
    public static class Utils
    {
        public static void GenerateBarcodeToStream(string content, BarcodeFormat format, Stream outputStream, int width = 300, int height = 150)
        {
            var barcodeWriter = new BarcodeWriter<SKBitmap>
            {
                Format = format,
                Renderer = new SKBitmapRenderer(),
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

        public static MemoryStream GetBarcode(string content, BarcodeFormat format, int width = 300, int height = 300)
        {
            var ms = new MemoryStream();
            GenerateBarcodeToStream(content, format, ms, width, height);
            return ms;
        }
    }
}
