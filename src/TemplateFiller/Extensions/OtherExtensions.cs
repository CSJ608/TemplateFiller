using System;
using System.IO;

namespace TemplateFiller.Extensions
{
    internal static class OtherExtensions
    {
        public static MemoryStream Copy(this Stream inputStream)
        {
            if (inputStream == null)
                throw new ArgumentNullException(nameof(inputStream));

            var memoryStream = new MemoryStream();

            // 重置流位置（如果支持搜索）
            if (inputStream.CanSeek && inputStream.Position != 0)
                inputStream.Position = 0;

            inputStream.CopyTo(memoryStream);
            memoryStream.Position = 0; // 重置内存流位置以便读取

            return memoryStream;
        }
    }
}
