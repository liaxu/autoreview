using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
                throw new ArgumentNullException("方法参数不能为空");
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
                throw new ArgumentNullException("方法参数不能为空");
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
            string destPath = Path.Combine(Path.GetDirectoryName(path), Path.GetFileNameWithoutExtension(path));
            if(Directory.Exists(destPath) == false)
            {
                Directory.CreateDirectory(destPath);
            }

            UnRAR(path, destPath);
            return destPath;
        }

        /// <summary>
        /// 暂时不用
        /// </summary>
        /// <param name="zipFilePath"></param>
        /// <param name="unzipDestPath"></param>
        private void UnZip(string zipFilePath,string unzipDestPath)
        {
            using (ZipInputStream s = new ZipInputStream(File.OpenRead(zipFilePath)))
            {
                ZipEntry theEntry;
                bool overWrite = false;
                string directoryName = "";
                if (!unzipDestPath.EndsWith("\\"))
                    unzipDestPath = unzipDestPath + "\\";
                while ((theEntry = s.GetNextEntry()) != null)
                {

                    string pathToZip = "";
                    pathToZip = theEntry.Name;
                    if (pathToZip != "")
                        directoryName = Path.GetDirectoryName(pathToZip) + "\\";
                    string fileName = Path.GetFileName(pathToZip);
                    Directory.CreateDirectory(unzipDestPath + directoryName);
                    if (fileName != "")
                    {
                        if ((File.Exists(unzipDestPath + directoryName + fileName) && overWrite) || (!File.Exists(unzipDestPath + directoryName + fileName)))
                        {
                            using (FileStream streamWriter = File.Create(unzipDestPath + directoryName + fileName))
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
            }

        }


        /// <summary>  
        /// 获取WinRAR.exe路径  
        /// </summary>  
        /// <returns>为空则表示未安装WinRAR</returns>  
        private string ExistsRAR()
        {
            RegistryKey regkey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WinRAR.exe");
            //RegistryKey regkey = Registry.ClassesRoot.OpenSubKey(@"Applications\WinRAR.exe\shell\open\command");  
            string strkey = regkey.GetValue("").ToString();
            regkey.Close();
            //return strkey.Substring(1, strkey.Length - 7);  
            return strkey;
        }

        /// <summary>  
        /// 解压RAR文件  
        /// </summary>  
        /// <param name="rarFilePath">要解压的文件路径</param>  
        /// <param name="unrarDestPath">解压路径（绝对路径）</param>  
        private void UnRAR(string rarFilePath, string unrarDestPath)
        {
            string rarexe = ExistsRAR();
            if (String.IsNullOrEmpty(rarexe))
            {
                throw new Exception("未安装WinRAR程序。");
            }
            try
            {
                //组合出需要shell的完整格式  
                string shellArguments = string.Format("x -t -ibck -o- \"{0}\" \"{1}\\\"", rarFilePath, unrarDestPath);

                //用Process调用  
                using (Process unrar = new Process())
                {
                    ProcessStartInfo startinfo = new ProcessStartInfo();
                    startinfo.FileName = rarexe;
                    startinfo.Arguments = shellArguments;               //设置命令参数  
                    startinfo.WindowStyle = ProcessWindowStyle.Hidden;  //隐藏 WinRAR 窗口  
                    unrar.StartInfo = startinfo;
                    unrar.Start();
                    unrar.WaitForExit();//等待解压完成  
                    unrar.Close();
                }
            }
            catch
            {
                throw;
            }
        }


    }
}
