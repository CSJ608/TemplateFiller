using System.IO;

namespace TemplateFiller.Utils
{
    internal static class Default
    {
        /// <summary>
        /// 打开写入文件的流，并在检测到目录不存在时创建目录
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static FileStream OpenWritableFileStream(string file)
        {
            var directory = Path.GetDirectoryName(file);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            return new FileStream(file, FileMode.Create, FileAccess.Write);
        }
    }
}
