using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateZip
{
    class Program
    {
        static void Main(string[] args)
        {
            var destFolder = "bin";
            var items = new List<string>() { "DANRAKU.dotm", "readme.txt", "Setup_WordAddin.vbs" };
            var zipFileName = "DANRAKU.zip";

            Directory.CreateDirectory(destFolder);
            items.ForEach(item => File.Copy(item, destFolder, true));
            ZipFile.CreateFromDirectory(destFolder, zipFileName);
        }
    }
}
