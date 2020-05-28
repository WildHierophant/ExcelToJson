using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using ExcelToJson.Properties;
using System.Diagnostics;

namespace ExcelToJson
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
            textBoxExcelURL.Text = Settings.Default.excelURL;
            textBoxJsonURL.Text = Settings.Default.jsonURL;
        }

        /// <summary>
        /// 文件列表，只读取xlsx格式与xls格式
        /// </summary>
        List<string> pathList = new List<string>();

        /// <summary>
        /// 表格数量统计
        /// </summary>
        int excelNum = 0;

        /// <summary>
        /// 成功导出为Json的表格数量统计
        /// </summary>
        int excelNumSuccessToJson = 0;

        /// <summary>
        /// 统计时间变量
        /// </summary>
        Stopwatch stopwatch = new Stopwatch();

        /// <summary>
        /// 总导表耗时统计
        /// </summary>
        double timeCount = 0f;

        /// <summary>
        /// ExcelURL按钮点击事件，即最上面那个...按钮，用于选择Excel文件所在文件夹地址
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonExcelURL_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = folderBrowserDialog.ShowDialog();
            if (dialogResult.ToString() == "OK")
            {
                textBoxExcelURL.Text = folderBrowserDialog.SelectedPath;
            }
        }

        /// <summary>
        /// JsonURL按钮点击事件，即最上面那个...按钮正下方的...按钮，用于选择输出Json文件的目标文件夹地址
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonJsonURL_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = folderBrowserDialog.ShowDialog();
            if (dialogResult.ToString() == "OK")
            {
                textBoxJsonURL.Text = folderBrowserDialog.SelectedPath;
            }
        }

        /// <summary>
        /// StartingExport按钮点击事件，即导出Json按钮，开始导出Json
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonStartingExport_Click(object sender, EventArgs e)
        {
            string excelFolderName = textBoxExcelURL.Text;
            string jsonFolderName = textBoxJsonURL.Text;
            getPath(excelFolderName, jsonFolderName);
            if (excelNum > 0)
            {
                for(int i = 1; i <= excelNum; i++)
                {
                    FileStream fileStream = new FileStream(pathList[i - 1], FileMode.Open, FileAccess.Read);
                    WriteToJson(fileStream);
                }
                listBox.Items.Add("表格总计:共读取到" + excelNum + "张表格");
                listBox.Items.Add("成功导出表格:共有" + excelNumSuccessToJson + "张表格成功导出为Json");
                listBox.Items.Add("错误表格:共有" + (excelNum - excelNumSuccessToJson) + "张表格导出失败");
                listBox.Items.Add("导出结束");
                listBox.Items.Add("耗时" + timeCount + "秒");
                listBox.Items.Add("---------------------------------------------------");
                excelNum = default;
                excelNumSuccessToJson = default;
                timeCount = default;
            }
        }

        /// <summary>
        /// ClearListBoxMessage按钮点击事件，清理ListBox中的消息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonClearListBoxMessage_Click(object sender, EventArgs e)
        {
            listBox.Items.Clear();
        }

        /// <summary>
        /// FormMain窗体关闭事件，关闭时记录两个textBox中的信息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            Settings.Default.excelURL = textBoxExcelURL.Text;
            Settings.Default.jsonURL = textBoxJsonURL.Text;
            Settings.Default.Save();
        }

        /// <summary>
        /// 获取目标文件夹内文件列表方法，只读取xlsx格式与xls格式
        /// </summary>
        /// <param name="excelPath"></param>
        /// <returns></returns>
        public List<string> getPath(string excelPath,string jsonPath)
        {
            //检查excelURL内地址是否为空
            if (excelPath != "")
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(excelPath);
                //检查excelURL内地址是否存在
                if (directoryInfo.Exists)
                {
                    FileInfo[] fileInfos = directoryInfo.GetFiles();
                    pathList.Clear();
                    foreach (FileInfo fileInfo in fileInfos)
                    {
                        if (fileInfo.Name.EndsWith("xlsx") || fileInfo.Name.EndsWith("xls"))
                        {
                            pathList.Add(fileInfo.FullName);
                        }
                    }
                }
                else
                {
                    listBox.Items.Add("Error:表格地址有误，请重新选择文件夹目录");
                    listBox.Items.Add("---------------------------------------------------");
                    pathList.Clear();
                    return pathList;
                }
            }
            else
            {
                listBox.Items.Add("Error:表格地址为空，请重新选择文件夹目录");
                listBox.Items.Add("---------------------------------------------------");
                pathList.Clear();
                return pathList;
            }
            //检查jsonURL内地址是否为空
            if (jsonPath != "")
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(jsonPath);
                //检查excelURL内地址是否存在
                if (directoryInfo.Exists)
                {
                    FileInfo[] fileInfos = directoryInfo.GetFiles();
                    foreach (FileInfo fileInfo in fileInfos)
                    {
                        if (fileInfo.Name.EndsWith("json"))
                        {
                            File.Delete(fileInfo.FullName);
                        }
                    }
                }
                else
                {
                    listBox.Items.Add("Error:导出地址有误，请重新选择文件夹目录");
                    listBox.Items.Add("---------------------------------------------------");
                    pathList.Clear();
                    return pathList;
                }
            }
            else
            {
                listBox.Items.Add("Error:导出地址为空，请重新选择文件夹目录");
                listBox.Items.Add("---------------------------------------------------");
                pathList.Clear();
                return pathList;
            }
            excelNum = pathList.Count;
            if (excelNum == 0)
            {
                listBox.Items.Add("Error:表格文件夹目录下未读取到表格，请重新选择文件夹");
                listBox.Items.Add("---------------------------------------------------");
            }
            else if (excelNum < 0)
            {
                listBox.Items.Add("Error:读取表格数量异常");
                listBox.Items.Add("---------------------------------------------------");
            }
            return pathList;
        }

        /// <summary>
        /// 将excel表格导出为全字符串的Json的方法，只读每个excel表中的第一张sheet表，表头在从第2行开始，数据从第4行开始
        /// </summary>
        /// <param name="workBook"></param>
        private void WriteToJson(FileStream fileStream)
        {
            stopwatch.Restart();
            //单次耗时
            double timer;
            XSSFWorkbook wookBook = new XSSFWorkbook(fileStream);
            //服务端表头所在行数，第2行
            var serverTitleRow = 1;
            //服务端数据开始行数，第4行
            var serverDataRow = 3;
            //服务端单元格类型行数，第3行
            var serverCellTypeRow = 2;
            //获取excel文件第一张sheet表格的表名
            string sheetName = wookBook.GetSheetName(0);
            //获取excel文件第一张sheet表格
            ISheet sheet = wookBook.GetSheetAt(0);
            //判断表头是否存在异常
            bool titleIsError = false;
            //Excel表格文件名
            string fileStreamName = fileStream.Name.Replace(textBoxExcelURL.Text + @"\", "");
            if (sheetName.StartsWith("Sheet") || sheetName.StartsWith("sheet") || sheetName == "")
            {
                stopwatch.Stop();
                timer = stopwatch.ElapsedMilliseconds * 0.001;
                timeCount += timer;
                listBox.Items.Add("Error:文件[" + fileStreamName + "]表格名称不符合规范，该表格无法导出Json，请修改");
                listBox.Items.Add("耗时" + timer + "秒");
                listBox.Items.Add("*********************************************");
            }
            else
            {
                //服务端表头
                IRow titleRow = sheet.GetRow(serverTitleRow);
                if (titleRow != null)
                {
                    //单元格类型表头
                    IRow cellTypeRow = sheet.GetRow(serverCellTypeRow);
                    //最后一格的编号，即列数
                    int columnCount = titleRow.LastCellNum;
                    //遍历表头是否有空值，如果有则不导出
                    for (int i = titleRow.FirstCellNum; i < columnCount; i++)
                    {
                        if (titleRow.GetCell(i) == null || titleRow.GetCell(i).ToString().Length == 0)
                        {
                            stopwatch.Stop();
                            timer = stopwatch.ElapsedMilliseconds * 0.001;
                            timeCount += timer;
                            listBox.Items.Add("Error:文件[" + fileStreamName + "]表格，第" + (i + 1) + "列表头存在数据异常，该表格无法导出Json，请修改");
                            listBox.Items.Add("耗时" + timer + "秒");
                            listBox.Items.Add("*********************************************");
                            titleIsError = true;
                            break;
                        }
                    }
                    if(titleIsError == false)
                    {
                        //遍历单元格类型表头是否有空值，如果有则不导出
                        for (int i = cellTypeRow.FirstCellNum; i < columnCount; i++)
                        {
                            if (cellTypeRow.GetCell(i) == null || cellTypeRow.GetCell(i).ToString().Length == 0)
                            {
                                stopwatch.Stop();
                                timer = stopwatch.ElapsedMilliseconds * 0.001;
                                timeCount += timer;
                                listBox.Items.Add("Error:文件[" + fileStreamName + "]表格，第" + (i + 1) + "列单元格类型存在数据异常，该表格无法导出Json，请修改");
                                listBox.Items.Add("耗时" + timer + "秒");
                                listBox.Items.Add("*********************************************");
                                titleIsError = true;
                                break;
                            }
                        }
                    }
                    try
                    {
                        if (titleIsError == false)
                        {
                            string jsonPath = textBoxJsonURL.Text + @"\" + sheetName + ".json";
                            FileStream fileStreamJson = new FileStream(jsonPath, FileMode.OpenOrCreate);
                            StreamWriter streamWriter = new StreamWriter(fileStreamJson);
                            streamWriter.Write("{\r\n");
                            streamWriter.Write("   \"" + sheetName.Trim() + "\":[\r\n");
                            //最后一行的编号
                            int rowCount = sheet.LastRowNum;
                            //遍历行
                            for (int i = serverDataRow; i <= rowCount; i++)
                            {
                                //获取行
                                IRow row = sheet.GetRow(i);
                                if (row == null)
                                {
                                    listBox.Items.Add("Error:文件[" + fileStreamName + "]表格，第" + (i + 1) + "行存在数据异常");
                                    break;
                                }
                                string strWrite = "        " + "{\r\n";
                                streamWriter.Write(strWrite);
                                string strWrite2 = "";
                                //遍历该行的列
                                for (int j = row.FirstCellNum; j < columnCount; j++)
                                {
                                    if (titleRow.GetCell(j) != null && titleRow.GetCell(j).ToString().Length != 0)
                                    {
                                        string value = "";
                                        if (row.GetCell(j) != null)
                                        {
                                            //目标单元格类型
                                            CellType cellType = row.GetCell(j).CellType;
                                            //单元格类型表头
                                            CellType rowCellType = default;
                                            switch (cellTypeRow.GetCell(j).ToString().Trim())
                                            {
                                                case "int":
                                                    rowCellType = CellType.Numeric;
                                                    break;
                                                case "string":
                                                    rowCellType = CellType.String;
                                                    break;
                                                case "bool":
                                                    rowCellType = CellType.Boolean;
                                                    break;
                                                case "":
                                                    rowCellType = CellType.Blank;
                                                    break;
                                            }
                                            if (cellType != rowCellType && cellType != CellType.Formula)
                                            {
                                                stopwatch.Stop();
                                                timer = stopwatch.ElapsedMilliseconds * 0.001;
                                                timeCount += timer;
                                                listBox.Items.Add("Error:文件[" + fileStreamName + "]表格，第" + (i + 1) + "行，第" + (j + 1) + "列单元格类型存在数据异常，该表格无法继续导出Json，请修改");
                                                listBox.Items.Add("耗时" + timer + "秒");
                                                listBox.Items.Add("*********************************************");
                                                streamWriter.Close();
                                                return;
                                            }
                                            else
                                            {
                                                switch (cellType)
                                                {
                                                    case CellType.Numeric:
                                                        value = row.GetCell(j).NumericCellValue.ToString().Trim();
                                                        break;
                                                    case CellType.String:
                                                        value = row.GetCell(j).StringCellValue.ToString().Trim();
                                                        break;
                                                    case CellType.Formula:
                                                        switch (rowCellType)
                                                        {
                                                            case CellType.Numeric:
                                                                value = row.GetCell(j).NumericCellValue.ToString().Trim();
                                                                break;
                                                            case CellType.String:
                                                                value = row.GetCell(j).StringCellValue.ToString().Trim();
                                                                break;
                                                            case CellType.Boolean:
                                                                value = row.GetCell(j).BooleanCellValue.ToString().Trim();
                                                                break;
                                                        }
                                                        break;
                                                    case CellType.Blank:
                                                        break;
                                                    case CellType.Boolean:
                                                        value = row.GetCell(j).BooleanCellValue.ToString().Trim();
                                                        break;
                                                    case CellType.Unknown:
                                                        listBox.Items.Add("Error:文件[" + fileStreamName + "]表格，第" + (i + 1) + "行，第" + (j + 1) + "列单元格类型未知");
                                                        break;
                                                    case CellType.Error:
                                                        listBox.Items.Add("Error:文件[" + fileStreamName + "]表格，第" + (i + 1) + "行，第" + (j + 1) + "列单元格类型存在异常");
                                                        break;
                                                }
                                            }
                                        }
                                        string title = titleRow.GetCell(j).ToString().Trim();
                                        strWrite2 += "          \"" + title + "\":\"" + value + "\",\r\n";
                                    }
                                }
                                int endStr = strWrite2.LastIndexOf(",");
                                if (endStr != -1)
                                {
                                    strWrite2 = strWrite2.Remove(endStr, 1);
                                    streamWriter.Write(strWrite2);
                                }
                                if (i == rowCount)
                                {
                                    streamWriter.Write("        }\r\n");
                                }
                                else
                                {
                                    streamWriter.Write("        },\r\n");
                                }
                            }
                            streamWriter.Write("   ]\r\n");
                            streamWriter.Write("}\r\n");
                            streamWriter.Close();
                            excelNumSuccessToJson += 1;
                            stopwatch.Stop();
                            timer = stopwatch.ElapsedMilliseconds * 0.001;
                            timeCount += timer;
                            listBox.Items.Add("文件[" + fileStreamName + "]导出Json成功");
                            listBox.Items.Add("耗时" + timer + "秒");
                            listBox.Items.Add("*********************************************");
                        }
                    }
                    catch (OverflowException)
                    {
                        stopwatch.Stop();
                        timer = stopwatch.ElapsedMilliseconds * 0.001;
                        timeCount += timer;
                        listBox.Items.Add("Error:堆栈溢出");
                        listBox.Items.Add("耗时" + timer + "秒");
                        listBox.Items.Add("*********************************************");
                        return;
                    }
                    catch (IOException ioException)
                    {
                        stopwatch.Stop();
                        timer = stopwatch.ElapsedMilliseconds * 0.001;
                        timeCount += timer;
                        listBox.Items.Add("Error:" + ioException.Message);
                        listBox.Items.Add("耗时" + timer + "秒");
                        listBox.Items.Add("*********************************************");
                        return;
                    }
                }
                else
                {
                    stopwatch.Stop();
                    timer = stopwatch.ElapsedMilliseconds * 0.001;
                    timeCount += timer;
                    listBox.Items.Add("Error:文件[" + fileStreamName + "]表格表头为空，该表格无法导出Json，请修改");
                    listBox.Items.Add("耗时" + timer + "秒");
                    listBox.Items.Add("*********************************************");
                }
            }
        }
    }
}
