using System;
using System.Collections.Generic;
using System.IO;
using AutoReview.Structure;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace AutoReview.ExcelParser
{
    public class ExcelParser : IExcelParser
    {
        private IWorkbook workbook = null;
        /// <summary>
        /// 根据课程分类获取该课程下的所有课程
        /// </summary>
        /// <param name="subjectName">课程分类名称</param>
        /// <returns>该分类下的所有课程</returns>
        public List<ClassWithScore> GetClassName(string subjectName)
        {
            if (string.IsNullOrEmpty(subjectName))
            {
                throw new Exception("subjectName参数不能为空");
            }
            if (this.workbook == null)
            {
                throw new Exception("文件初始化失败");
            }
            List<ClassWithScore> result = new List<ClassWithScore>();
            ISheet sheet = workbook.GetSheet(subjectName);
            if (sheet == null)
            {
                throw new Exception("没有找到" + subjectName);
            }
            IRow row = null;
            ICell cell = null;
            string cellValue = "";
            ClassWithScore classWithScore = null;
            for (var i = 6; i <= sheet.LastRowNum; i++)
            {

                row = sheet.GetRow(i);
                if (row != null)
                {
                    for (var j = 0; j < 6; j += 2)
                    {
                        cell = row.GetCell(j);
                        if (cell != null)
                        {
                            cellValue = GetCellValue(cell);
                            if (string.IsNullOrEmpty(cellValue))
                            {
                                continue;
                            }
                            classWithScore = new ClassWithScore();
                            classWithScore.ClassName = cellValue;
                        }
                        cell = row.GetCell(j + 1);
                        if (cell != null)
                        {
                            cellValue = GetCellValue(cell);
                            if (string.IsNullOrEmpty(cellValue))
                            {
                                continue;
                            }
                            classWithScore.Score = int.Parse(cellValue);
                        }
                        if (classWithScore != null) result.Add(classWithScore);
                        classWithScore = null;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 获取所有强支撑课程
        /// </summary>
        /// <returns>强支撑课程list</returns>
        public List<StrongSupportClass> GetHighSupportClass()
        {
            if (this.workbook == null)
            {
                throw new Exception("文件初始化失败");
            }
            ISheet sheet = this.workbook.GetSheetAt(0);
            if (sheet == null)
            {
                throw new Exception("没有找到Sheet");
            }
            IRow row = null;
            IRow requireRow = null;
            ICell cell = null;
            string cellValue = "";
            List<StrongSupportClass> result = new List<StrongSupportClass>();
            StrongSupportClass strongSupportClass = null;
            for (var i = 3; i <= sheet.LastRowNum; i++)
            {
                strongSupportClass = new StrongSupportClass();
                strongSupportClass.supportPoint = new List<string>();
                row = sheet.GetRow(i);
                if (row != null)
                {
                    requireRow = sheet.GetRow(2);
                    for (var j = 1; j <= row.LastCellNum; j++)
                    {
                        cell = row.GetCell(j);
                        if (cell != null && "H".Equals(cellValue = GetCellValue(cell).ToUpper()))
                        {
                            if (string.IsNullOrEmpty(strongSupportClass.ClassName))
                            {
                                strongSupportClass.ClassName = GetCellValue(row.GetCell(0));
                            }                           
                            cell = requireRow.GetCell(j);
                            if(!string.IsNullOrEmpty(cellValue = GetCellValue(cell)))
                            {
                                strongSupportClass.supportPoint.Add(cellValue);
                            }
                        }
                    }
                    if (strongSupportClass.supportPoint.Count > 0)
                    {
                        result.Add(strongSupportClass);
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 获取核心专业课
        /// </summary>
        /// <returns>核心专业课列表</returns>
        public List<string> GetProfessinalCoreClass()
        {
            if (this.workbook == null)
            {
                throw new Exception("文件初始化失败");
            }
            List<string> result = new List<string>();
            ISheet sheet = workbook.GetSheet("专业核心课");
            if (sheet != null)
            {
                if (sheet.LastRowNum >= 3)
                {
                    IRow row = null;
                    ICell cell = null;
                    for (var i = 3; i <= sheet.LastRowNum; i++)
                    {
                        row = sheet.GetRow(i);
                        if (row != null)
                        {
                            cell = row.GetCell(0);
                            if (cell != null && !string.IsNullOrEmpty(GetCellValue(cell)))
                            {
                                result.Add(GetCellValue(cell));
                            }
                        }
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 根据课程分类获取该课程分类的总分
        /// </summary>
        /// <param name="subjectName">课程分类名称</param>
        /// <returns>总分</returns>
        public int GetScore(string subjectName)
        {
            if (this.workbook == null)
            {
                throw new Exception("文件初始化失败");
            }
            ISheet sheet = this.workbook.GetSheet(subjectName);
            if (sheet == null)
            {
                throw new Exception("没有找到" + subjectName);
            }
            int sum = 0;
            int tmp = 0;
            IRow row = null;
            ICell cell = null;
            string cellValue = "";
            if (sheet != null)
            {

                for (var i = 4; i <= sheet.LastRowNum; i++)
                {
                    if (i == 5) continue;
                    row = sheet.GetRow(i);
                    if (i == 4)
                    {
                        cell = row.GetCell(5);

                        if (cell != null && !string.IsNullOrEmpty(cellValue = GetCellValue(cell)))
                        {
                            tmp = int.Parse(cellValue);
                            sum += tmp;
                        }
                    }
                    else
                    {
                        cell = row.GetCell(1);

                        if (cell != null && !string.IsNullOrEmpty(cellValue = GetCellValue(cell)))
                        {
                            tmp = int.Parse(cellValue);
                            sum += tmp;
                        }
                        cell = row.GetCell(3);
                        if (cell != null && !string.IsNullOrEmpty(cellValue = GetCellValue(cell)))
                        {
                            tmp = int.Parse(cellValue);
                            sum += tmp;
                        }
                    }
                }
            }
            return sum;
        }

        /// <summary>
        /// 获取总分
        /// </summary>
        /// <returns>总分</returns>
        public int GetTotalScore()
        {
            if (this.workbook == null)
            {
                throw new Exception("文件初始化失败");
            }
            ISheet sheet = this.workbook.GetSheet("学分统计");
            if (sheet == null)
            {
                return 0;
            }
            IRow row = sheet.GetRow(5);
            if (row != null)
            {
                ICell cell = row.GetCell(1);
                if (cell != null)
                {
                    var cellValue = GetCellValue(cell);
                    return int.Parse(string.IsNullOrEmpty(cellValue) ? "0" : cellValue);
                }
            }
            return 0;
        }

        /// <summary>
        /// 初始化文件
        /// </summary>
        /// <param name="path">初始化的文件</param>
        public void Init(string path)
        {
            Stream fileStream = OpenClasspathResource(path);
            if (string.IsNullOrEmpty(path))
            {
                throw new ArgumentNullException("path参数不能为空");
            }
            if (".xlsx".Equals(Path.GetExtension(path).ToLower()))
            {
                workbook = new XSSFWorkbook(fileStream);
            }
            else if (".xls".Equals(Path.GetExtension(path).ToLower()))
            {
                workbook = new HSSFWorkbook(fileStream);
            }
            else
            {
                throw new Exception("文件不是有效的Excel文件");
            }
            fileStream.Close();
        }

        /// <summary>
        /// 将文件转换成文件流
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private Stream OpenClasspathResource(String fileName)
        {
            FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            return file;
        }

        /// <summary>
        /// 获取单元格的值
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns></returns>
        private string GetCellValue(ICell cell)
        {
            string value = "";
            switch (cell.CellType)
            {
                case CellType.Blank:
                    value = "";
                    break;
                case CellType.Numeric:
                    value = cell.NumericCellValue.ToString();
                    break;
                case CellType.String:
                    value = cell.StringCellValue;
                    break;
                default:
                    value = "";
                    break;
            }
            return value;
        }

        /// <summary>
        /// 释放对象
        /// </summary>
        public void Dispose()
        {
            if (this.workbook != null)
            {
               this.workbook.Close();
            }
        }
    }
}
