using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReview.ZipFileManager
{
    interface IZipFileManager
    {
        string UnZip(string path);
        void DeleteFolder(string path);
        List<string> FindFile(string name);
    }
}
