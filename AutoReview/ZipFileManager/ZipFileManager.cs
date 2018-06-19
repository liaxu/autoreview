using ICSharpCode.SharpZipLib.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace AutoReview.ZipFileManager
{
    /// <summary>
    /// Zip文件相关操作：解压、删除目录及子文件、根据文件名称查找文件
    /// </summary>
    public class ZipFileManager : IZipFileManager
    {
        /// <summary>
        /// 删除某文件夹及其子目录下的文件
        /// </summary>
        /// <param name="path">文件夹路径</param>
        public void DeleteFolder(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                throw new ArgumentNullException("path不能为空");
            }
            if (Directory.Exists(path))
            {
                try
                {
                    DirectoryInfo dir = new DirectoryInfo(path);
                    FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();  //返回目录中所有文件和子目录
                    foreach (FileSystemInfo i in fileinfo)
                    {
                        if (i is DirectoryInfo)            //判断是否文件夹
                        {
                            DirectoryInfo subdir = new DirectoryInfo(i.FullName);
                            subdir.Delete(true);          //删除子目录和文件
                        }
                        else
                        {
                            File.Delete(i.FullName);      //删除指定文件
                        }
                    }
                    Directory.Delete(path);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

        }

        /// <summary>
        /// 查找带有某个名称特征的文件
        /// </summary>
        /// <param name="path">查找的路径</param>
        /// <param name="name">所查找的文件名称关键字</param>
        /// <returns></returns>
        public List<string> FindFile(string path, string name)
        {
            if (string.IsNullOrEmpty(path) || string.IsNullOrEmpty(name))
            {
                throw new ArgumentNullException("方法参数不能为为空");
            }
            string[] fileDir = Directory.GetFiles(path, "*" + name + "*", SearchOption.AllDirectories);
            return new List<string>(fileDir);
        }


        /// <summary>
        /// 解压文件
        /// </summary>
        /// <param name="path">压缩文件的完整路径</param>
        /// <returns></returns>
        public string UnZip(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentNullException("方法参数不能为空");
            if (!".zip".Equals(Path.GetExtension(path).ToLower()))
            {
                throw new Exception("文件类型错误");
            }
            using (ZipInputStream s = new ZipInputStream(File.OpenRead(path)))
            {
                ZipEntry theEntry;
                bool overWrite = false;
                string tarDirectory = "";
                string directoryName = "";
                tarDirectory = Path.GetDirectoryName(path);
                if (!tarDirectory.EndsWith("\\"))
                    tarDirectory = tarDirectory + "\\";
                while ((theEntry = s.GetNextEntry()) != null)
                {

                    string pathToZip = "";
                    pathToZip = theEntry.Name;
                    if (pathToZip != "")
                        directoryName = Path.GetDirectoryName(pathToZip) + "\\";
                    string fileName = Path.GetFileName(pathToZip);
                    Directory.CreateDirectory(tarDirectory + directoryName);
                    if (fileName != "")
                    {
                        if ((File.Exists(tarDirectory + directoryName + fileName) && overWrite) || (!File.Exists(tarDirectory + directoryName + fileName)))
                        {
                            using (FileStream streamWriter = File.Create(tarDirectory + directoryName + fileName))
                            {
                                int size = 2048;
                                byte[] data = new byte[2048];
                                while (true)
                                {
                                    size = s.Read(data, 0, data.Length);

                                    if (size > 0)
                                        streamWriter.Write(data, 0, size);
                                    else
                                        break;
                                }
                                streamWriter.Close();
                            }
                        }
                    }
                }

                s.Close();
                return tarDirectory + directoryName;
            }

        }
    }
}
