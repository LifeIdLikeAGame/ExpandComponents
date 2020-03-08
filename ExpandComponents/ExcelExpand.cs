using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace ExpandComponents
{
    /********************************************************************************

    ** 类名称： ExcelExpand

    ** 描述：导入导出Excel拓展

    ** 引用： NPOI.dll   NPOI.OOML.dll  NPOI.OpenXml4Net.dll  NPOI.OpenXmlFormats.dll  ICSharpCode.SharpZipLib.dll(版本：1.0.0.999)

    ** 作者： LW

    *********************************************************************************/

    /// <summary>
    /// 导入导出Excel拓展
    /// </summary>
    public static class ExcelExpand
    {
        #region 导入读取文件拓展

        #region Microsoft.ACE.OLEDB.12.0 导入

        /// <summary>
        /// 将Excel数据表格转换为DateTable,返回excel第一个有数据的表
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <param name="del">是否删除原文件</param>
        /// <returns></returns>
        public static DataTable ExcelToDataTable(string path, bool del = false)
        {
            bool flag = !File.Exists(path);
            if (flag)
            {
                throw new Exception("未找到 path 中指定的文件。");
            }
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            OleDbConnection oleDbConnection = new OleDbConnection(connectionString);
            DataSet dataSet = new DataSet();
            DataTable result;
            try
            {
                oleDbConnection.Open();
                DataTable oleDbSchemaTable = oleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                DataTable dataTable = new DataTable();
                for (int i = 0; i < oleDbSchemaTable.Rows.Count; i++)
                {
                    string value = oleDbSchemaTable.Rows[i]["TABLE_NAME"].ToString();
                    bool flag2 = string.IsNullOrEmpty(value);
                    if (!flag2)
                    {
                        string selectCmddText = "select * from [" + value + "]";
                        OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter(selectCmddText, oleDbConnection);
                        oleDbDataAdapter.Fill(dataSet, "excelData");
                        var count = dataSet.Tables[0].Rows.Count;
                        if (count > 0) { dataTable = dataSet.Tables[0]; break; };
                    }
                }
                for (int i = dataTable.Rows.Count - 1; i >= 0; i--)
                {
                    int num = 0;
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        string value = dataTable.Rows[i][j].ToString();
                        bool flag2 = string.IsNullOrEmpty(value);
                        if (flag2)
                        {
                            num++;
                        }
                    }
                    bool flag3 = num == dataTable.Columns.Count;
                    if (flag3)
                    {
                        dataTable.Rows.RemoveAt(i);
                    }
                }
                result = dataTable;
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                bool flag4 = oleDbConnection.State > ConnectionState.Closed;
                if (flag4)
                {
                    oleDbConnection.Close();
                }
                oleDbConnection.Dispose();
                if (del)
                {
                    if (File.Exists(path)) File.Delete(path);
                }
            }
            return result;
        }
        /// <summary>
        /// 将Excel数据表格转换为DateTable，指定表名
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <param name="TableName">Excel表名</param>
        /// <param name="del">是否删除原文件</param>
        /// <returns></returns>
        public static DataTable ExcelToDataTable(string path, string TableName, bool del = false)
        {
            bool flag = !File.Exists(path);
            if (flag)
            {
                throw new Exception("未找到 path 中指定的文件。");
            }
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            OleDbConnection oleDbConnection = new OleDbConnection(connectionString);
            DataSet dataSet = new DataSet();
            DataTable result;
            try
            {
                oleDbConnection.Open();
                DataTable oleDbSchemaTable = oleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                DataTable dataTable = new DataTable();
                string selectCmddText = "select * from [" + TableName + "]";
                OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter(selectCmddText, oleDbConnection);
                oleDbDataAdapter.Fill(dataSet, "excelData");
                var count = dataSet.Tables[0].Rows.Count;
                if (count > 0) { dataTable = dataSet.Tables[0]; }
                for (int i = dataTable.Rows.Count - 1; i >= 0; i--)
                {
                    int num = 0;
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        string value = dataTable.Rows[i][j].ToString();
                        bool flag2 = string.IsNullOrEmpty(value);
                        if (flag2)
                        {
                            num++;
                        }
                    }
                    bool flag3 = num == dataTable.Columns.Count;
                    if (flag3)
                    {
                        dataTable.Rows.RemoveAt(i);
                    }
                }
                result = dataTable;
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                bool flag4 = oleDbConnection.State > ConnectionState.Closed;
                if (flag4)
                {
                    oleDbConnection.Close();
                }
                oleDbConnection.Dispose();
                if (del)
                {
                    if (File.Exists(path)) File.Delete(path);
                }
            }
            return result;
        }

        #endregion


        #region NPOI 导入

        #region 将 Excel 文件读取到 DataTable
        /// <summary>
        /// 将 Excel 文件读取到 <see cref="DataTable"/>
        /// </summary>
        /// <param name="filePath">文件完整路径名</param>
        /// <param name="sheetName">指定读取 Excel 工作薄 sheet 的名称</param>
        /// <param name="firstRowIsColumnName">首行是否为 <see cref="DataColumn.ColumnName"/></param>
        /// <returns><see cref="DataTable"/>数据表</returns>
        public static DataTable ReadExcelToDataTable(string filePath, string sheetName = null, bool firstRowIsColumnName = true)
        {
            if (string.IsNullOrEmpty(filePath)) return null;
            if (!File.Exists(filePath)) return null;

            //根据指定路径读取文件
            var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            if (fileStream == null || fileStream.Length <= 0) return null;

            //定义要返回的datatable对象
            DataTable data = new DataTable();
            //Excel行号
            var RowIndex = 0;
            //excel工作表
            ISheet sheet = null;

            //数据开始行(排除标题行)
            int startRow = 0;
            try
            {
                //根据文件流创建excel数据结构
                IWorkbook workbook = WorkbookFactory.Create(fileStream);


                if (string.IsNullOrEmpty(sheetName)) sheet = workbook.GetSheetAt(0);
                else
                {
                    sheet = workbook.GetSheet(sheetName);

                    //如果没有找到指定的sheetName对应的sheet，则获取第一个sheet
                    if (sheet == null) sheet = workbook.GetSheetAt(0);
                }

                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    if (firstRow == null) new Exception("未获取到表头数据");
                    RowIndex = firstRow.RowNum + 1;
                    //一行最后一个cell的编号 即总的列数
                    int cellCount = firstRow.LastCellNum;

                    //如果第一行是标题列名
                    if (firstRowIsColumnName)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    cellValue = cellValue.Trim().Replace(" ", "");
                                    if (data.Columns.Contains(cellValue)) throw new Exception($"表头存在多个列名为“{cellValue}”的列！");
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                                else
                                {
                                    DataColumn column = new DataColumn("Column" + (i + 1));
                                    data.Columns.Add(column);
                                }
                            }
                            else
                            {
                                DataColumn column = new DataColumn("Column" + (i + 1));
                                data.Columns.Add(column);
                            }
                        }
                        if (cellCount > 0)
                        {
                            DataColumn column = new DataColumn("RowNum");
                            data.Columns.Add(column);
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }

                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; // 没有数据的行默认是 null，如果为 null 则不添加
                        var blankCount = 0;//空白单元格数
                        DataRow dataRow = data.NewRow();
                        RowIndex = row.RowNum + 1;
                        for (int j = row.FirstCellNum; j < row.FirstCellNum + cellCount; ++j)
                        {
                            //同理，没有数据的单元格都默认是null
                            ICell cell = row.GetCell(j);
                            //判断单元格是否为空白
                            if (cell == null || cell.CellType == CellType.Blank) { blankCount++; continue; }
                            if (cell != null)
                            {
                                if (cell.CellType == CellType.Numeric)
                                {
                                    //判断是否日期类型
                                    if (DateUtil.IsCellDateFormatted(cell))
                                    {
                                        dataRow[j] = row.GetCell(j).DateCellValue;
                                    }
                                    else
                                    {
                                        dataRow[j] = row.GetCell(j).ToString().Trim();
                                    }
                                }
                                else if (cell.CellType == CellType.Formula)
                                {
                                    CellType CellType = row.GetCell(j).CachedFormulaResultType;
                                    //判断是否公式计算值
                                    if (CellType == CellType.String)
                                        dataRow[j] = row.GetCell(j).StringCellValue.ToString().Trim();
                                    else if (CellType == CellType.Numeric)
                                        dataRow[j] = row.GetCell(j).NumericCellValue.ToString().Trim();
                                    else if (CellType == CellType.Blank) dataRow[j] = "";
                                    else dataRow[j] = row.GetCell(j).ToString().Trim();
                                }
                                else
                                {
                                    dataRow[j] = row.GetCell(j).ToString().Trim();
                                }
                            }
                        }
                        if (blankCount == cellCount) continue;
                        if (cellCount > 0)
                        {
                            var ColumnIndex = data.Columns.Count - 1;
                            dataRow[ColumnIndex] = RowIndex.ToString();
                        }
                        data.Rows.Add(dataRow);
                    }
                }
                return data;
            }
            catch (Exception e)
            {
                var err = new Exception($"行号:{RowIndex},错误：{e.Message}");
                throw err;
            }
            finally
            {
                //fileStream.Flush();
                fileStream.Close();
                fileStream.Dispose();
            }
        }
        /// <summary>
        /// 将 <see cref="Stream"/> 对象读取到 <see cref="DataTable"/>,指定表头索引
        /// </summary>
        /// <param name="stream">当前 <see cref="Stream"/> 对象</param>
        /// <param name="sheetName">指定读取 Excel 工作薄 sheet 的名称</param>
        /// <param name="HeaderIndex">指定表头行索引,默认第一行索引为:0</param>
        /// <returns><see cref="DataTable"/>数据表</returns>
        public static DataTable ReadExcelToDataTable_Specifies_Header(string filePath, string sheetName = null, int HeaderIndex = 0)
        {
            if (string.IsNullOrEmpty(filePath)) return null;
            if (!File.Exists(filePath)) return null;

            //根据指定路径读取文件
            var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            if (fileStream == null || fileStream.Length <= 0) return null;

            //定义要返回的datatable对象
            DataTable data = new DataTable();
            //Excel行号
            var RowIndex = 1;
            //excel工作表
            ISheet sheet = null;

            //数据开始行(排除标题行)
            int startRow = 0;
            try
            {
                //根据文件流创建excel数据结构
                IWorkbook workbook = WorkbookFactory.Create(fileStream);


                if (string.IsNullOrEmpty(sheetName)) sheet = workbook.GetSheetAt(0);
                else
                {
                    sheet = workbook.GetSheet(sheetName);

                    //如果没有找到指定的sheetName对应的sheet，则获取第一个sheet
                    if (sheet == null) sheet = workbook.GetSheetAt(0);
                }

                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(HeaderIndex);
                    if (firstRow == null || firstRow.FirstCellNum < 0) new Exception("未获取到表头数据");
                    RowIndex = firstRow.RowNum + 1;
                    //一行最后一个cell的编号 即总的列数
                    int cellCount = firstRow.LastCellNum;

                    //如果第一行是标题列名
                    if (HeaderIndex >= 0)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    cellValue = cellValue.Trim().Replace(" ", "");
                                    if (data.Columns.Contains(cellValue)) throw new Exception($"表头存在多个列名为“{cellValue}”的列！");
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                                else
                                {
                                    DataColumn column = new DataColumn("Column" + (i + 1));
                                    data.Columns.Add(column);
                                }
                            }
                            else
                            {
                                DataColumn column = new DataColumn("Column" + (i + 1));
                                data.Columns.Add(column);
                            }
                        }
                        if (cellCount > 0)
                        {
                            DataColumn column = new DataColumn("RowNum");
                            data.Columns.Add(column);
                        }
                        startRow = HeaderIndex + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }

                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null || row.FirstCellNum < 0) continue; // 没有数据的行默认是 null，如果为 null 则不添加
                        RowIndex = row.RowNum + 1;
                        var blankCount = 0;//空白单元格数
                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < row.FirstCellNum + cellCount; ++j)
                        {
                            //同理，没有数据的单元格都默认是null
                            ICell cell = row.GetCell(j);
                            //判断单元格是否为空白
                            if (cell == null || cell.CellType == CellType.Blank) { blankCount++; continue; }
                            if (cell != null)
                            {
                                if (cell.CellType == CellType.Numeric)
                                {
                                    //判断是否日期类型
                                    if (DateUtil.IsCellDateFormatted(cell))
                                    {
                                        dataRow[j] = row.GetCell(j).DateCellValue;
                                    }
                                    else
                                    {
                                        dataRow[j] = row.GetCell(j).ToString().Trim();
                                    }
                                }
                                else if (cell.CellType == CellType.Formula)
                                {
                                    CellType CellType = row.GetCell(j).CachedFormulaResultType;
                                    //判断是否公式计算值
                                    if (CellType == CellType.String)
                                        dataRow[j] = row.GetCell(j).StringCellValue.ToString().Trim();
                                    else if (CellType == CellType.Numeric)
                                        dataRow[j] = row.GetCell(j).NumericCellValue.ToString().Trim();
                                    else if (CellType == CellType.Blank) dataRow[j] = "";
                                    else dataRow[j] = row.GetCell(j).ToString().Trim();
                                }
                                else
                                {
                                    dataRow[j] = row.GetCell(j).ToString().Trim();
                                }
                            }
                        }
                        if (blankCount == cellCount) continue;
                        if (cellCount > 0)
                        {
                            var ColumnIndex = data.Columns.Count - 1;
                            dataRow[ColumnIndex] = RowIndex.ToString();
                        }
                        data.Rows.Add(dataRow);
                    }
                }
                return data;
            }
            catch (Exception e)
            {
                var err = new Exception($"行号:{RowIndex},错误：{e.Message}");
                throw err;
            }
            finally
            {
                fileStream.Close();
                fileStream.Dispose();
            }
        }
        #endregion

        #region 将 Stream 对象读取到 DataTable
        /// <summary>
        /// 将 <see cref="Stream"/> 对象读取到 <see cref="DataTable"/>
        /// </summary>
        /// <param name="stream">当前 <see cref="Stream"/> 对象</param>
        /// <param name="sheetName">指定读取 Excel 工作薄 sheet 的名称</param>
        /// <param name="firstRowIsColumnName">首行是否为 <see cref="DataColumn.ColumnName"/></param>
        /// <returns><see cref="DataTable"/>数据表</returns>
        public static DataTable ReadStreamToDataTable(Stream stream, string sheetName = null, bool firstRowIsColumnName = true)
        {
            if (stream == null || stream.Length <= 0) return null;

            //定义要返回的datatable对象
            var data = new DataTable();

            //excel工作表
            ISheet sheet = null;
            //Excel行号
            var RowIndex = 1;
            //数据开始行(排除标题行)
            int startRow = 0;
            try
            {
                //根据文件流创建excel数据结构,NPOI的工厂类WorkbookFactory会自动识别excel版本，创建出不同的excel数据结构
                IWorkbook workbook = WorkbookFactory.Create(stream);

                //如果有指定工作表名称
                if (!string.IsNullOrEmpty(sheetName))
                {
                    sheet = workbook.GetSheet(sheetName);
                    //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    if (sheet == null)
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                }
                else
                {
                    //如果没有指定的sheetName，则尝试获取第一个sheet
                    sheet = workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    if (firstRow == null || firstRow.FirstCellNum < 0)throw new Exception("未获取到表头数据");
                    RowIndex = firstRow.RowNum + 1;
                    //一行最后一个cell的编号 即总的列数
                    int cellCount = firstRow.LastCellNum;

                    //如果第一行是标题列名
                    if (firstRowIsColumnName)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    cellValue = cellValue.Trim().Replace(" ", "");
                                    if (data.Columns.Contains(cellValue)) throw new Exception($"表头存在多个列名为“{cellValue}”的列！");
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                                else
                                {
                                    DataColumn column = new DataColumn("Column" + (i + 1));
                                    data.Columns.Add(column);
                                }
                            }
                            else
                            {
                                DataColumn column = new DataColumn("Column" + (i + 1));
                                data.Columns.Add(column);
                            }
                        }
                        if (cellCount > 0)
                        {
                            DataColumn column = new DataColumn("RowNum");
                            data.Columns.Add(column);
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }

                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null || row.FirstCellNum < 0) continue; //没有数据的行默认是null　
                        RowIndex = row.RowNum + 1;
                        var blankCount = 0;//空白单元格数
                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < row.FirstCellNum + cellCount; ++j)
                        {
                            //同理，没有数据的单元格都默认是null
                            ICell cell = row.GetCell(j);
                            //判断单元格是否为空白
                            if (cell == null || cell.CellType == CellType.Blank) { blankCount++; continue; }
                            if (cell != null)
                            {
                                if (cell.CellType == CellType.Numeric)
                                {
                                    //判断是否日期类型
                                    if (DateUtil.IsCellDateFormatted(cell))
                                    {
                                        dataRow[j] = row.GetCell(j).DateCellValue;
                                    }
                                    else
                                    {
                                        dataRow[j] = row.GetCell(j).ToString().Trim();
                                    }
                                }
                                else if (cell.CellType == CellType.Formula)
                                {
                                    CellType CellType = row.GetCell(j).CachedFormulaResultType;
                                    //判断是否公式计算值
                                    if(CellType == CellType.String)
                                        dataRow[j] = row.GetCell(j).StringCellValue.ToString().Trim();
                                    else if (CellType == CellType.Numeric)
                                        dataRow[j] = row.GetCell(j).NumericCellValue.ToString().Trim();
                                    else if (CellType == CellType.Blank) dataRow[j] = "";
                                    else dataRow[j] = row.GetCell(j).ToString().Trim();
                                }
                                else
                                {
                                    dataRow[j] = row.GetCell(j).ToString().Trim();
                                }
                            }
                        }
                        if (blankCount == cellCount) continue;
                        if (cellCount > 0)
                        {
                            var ColumnIndex = data.Columns.Count - 1;
                            dataRow[ColumnIndex] = RowIndex.ToString();
                        }
                        data.Rows.Add(dataRow);
                    }
                }
                return data;
            }
            catch (Exception e)
            {
                var err = new Exception($"行号:{RowIndex},错误：{e.Message}");
                throw err;
            }
            finally
            {
                stream.Close(); // 关闭流
                stream.Dispose();
            }
        }
        /// <summary>
        /// 将 <see cref="Stream"/> 对象读取到 <see cref="DataTable"/>,指定表头索引,含 RowNum 行号
        /// </summary>
        /// <param name="stream">当前 <see cref="Stream"/> 对象</param>
        /// <param name="sheetName">指定读取 Excel 工作薄 sheet 的名称</param>
        /// <param name="HeaderIndex">指定表头行索引,默认第一行索引为:0</param>
        /// <returns><see cref="DataTable"/>数据表</returns>
        public static DataTable ReadStreamToDataTable_Specifies_Header(Stream stream, string sheetName = null, int HeaderIndex = 0)
        {
            if (stream == null || stream.Length <= 0) return null;

            //定义要返回的datatable对象
            var data = new DataTable();

            //excel工作表
            ISheet sheet = null;
            //Excel行号
            var RowIndex = 1;
            //数据开始行(排除标题行)
            int startRow = 0;
            try
            {
                //根据文件流创建excel数据结构,NPOI的工厂类WorkbookFactory会自动识别excel版本，创建出不同的excel数据结构
                IWorkbook workbook = WorkbookFactory.Create(stream);

                //如果有指定工作表名称
                if (!string.IsNullOrEmpty(sheetName))
                {
                    sheet = workbook.GetSheet(sheetName);
                    //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    if (sheet == null)
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                }
                else
                {
                    //如果没有指定的sheetName，则尝试获取第一个sheet
                    sheet = workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(HeaderIndex);
                    if (firstRow == null || firstRow.FirstCellNum < 0) new Exception("未获取到表头数据");
                    RowIndex = firstRow.RowNum + 1;
                    //一行最后一个cell的编号 即总的列数
                    int cellCount = firstRow.LastCellNum;

                    //如果第一行是标题列名
                    if (HeaderIndex >= 0)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    cellValue = cellValue.Trim().Replace(" ", "");
                                    if (data.Columns.Contains(cellValue)) throw new Exception($"表头存在多个列名为“{cellValue}”的列！");
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                                else
                                {
                                    DataColumn column = new DataColumn("Column" + (i + 1));
                                    data.Columns.Add(column);
                                }
                            }
                            else
                            {
                                DataColumn column = new DataColumn("Column" + (i + 1));
                                data.Columns.Add(column);
                            }
                        }
                        if (cellCount > 0)
                        {
                            DataColumn column = new DataColumn("RowNum");
                            data.Columns.Add(column);
                        }
                        startRow = HeaderIndex + 1;
                    }

                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null || row.FirstCellNum < 0) continue; //没有数据的行默认是null　　
                        RowIndex = row.RowNum + 1;
                        var blankCount = 0;//空白单元格数
                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < row.FirstCellNum + cellCount; ++j)
                        {
                            //同理，没有数据的单元格都默认是null
                            ICell cell = row.GetCell(j);
                            //判断单元格是否为空白
                            if (cell == null || cell.CellType == CellType.Blank) { blankCount++; continue; }
                            if (cell != null)
                            {
                                if (cell.CellType == CellType.Numeric)
                                {
                                    //判断是否日期类型
                                    if (DateUtil.IsCellDateFormatted(cell))
                                    {
                                        dataRow[j] = row.GetCell(j).DateCellValue;
                                    }
                                    else
                                    {
                                        dataRow[j] = row.GetCell(j).ToString().Trim();
                                    }
                                }
                                else if (cell.CellType == CellType.Formula)
                                {
                                    CellType CellType = row.GetCell(j).CachedFormulaResultType;
                                    //判断是否公式计算值
                                    if (CellType == CellType.String)
                                        dataRow[j] = row.GetCell(j).StringCellValue.ToString().Trim();
                                    else if (CellType == CellType.Numeric)
                                        dataRow[j] = row.GetCell(j).NumericCellValue.ToString().Trim();
                                    else if (CellType == CellType.Blank) dataRow[j] = "";
                                    else dataRow[j] = row.GetCell(j).ToString().Trim();
                                }
                                else
                                {
                                    dataRow[j] = row.GetCell(j).ToString().Trim();
                                }
                            }
                        }
                        if (blankCount == cellCount) continue;
                        if (RowIndex >= 61)
                        {

                        }
                        if (cellCount > 0)//行号
                        {
                            var ColumnIndex = data.Columns.Count - 1;
                            dataRow[ColumnIndex] = RowIndex.ToString();
                        }
                        data.Rows.Add(dataRow);
                    }
                }
                return data;
            }
            catch (Exception e)
            {
                var err = new Exception($"行号:{RowIndex},错误：{e.Message}");
                throw err;
            }
            finally
            {
                stream.Close(); // 关闭流
                stream.Dispose();
            }
        }
        //判断某行某列有问题
        private static int CheckRowError(HSSFCell cell)
        {
            //判断各个单元格是否为空
            if (cell == null || cell.Equals(""))
            {
                return -1;
            }
            return 0;
        }
        #endregion

        /// <summary>
        /// 是否为Excel文件
        /// </summary>
        /// <returns></returns>
        public static bool FileIsExcel(string path)
        {
            var _fileInfo = new FileInfo(path);
            if (_fileInfo == null) return false;
            var ext = _fileInfo.Extension.ToLower();
            if (ext == ".xls" || ext == ".xlsx") return true;
            return false;
        }

        #endregion

        #endregion

        #region 导出Excel拓展

        #region NPOI 导出 单元格拓展
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列,第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, ref int Index, string Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (!string.IsNullOrEmpty(Value)) cell.SetCellValue(Value);
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, ref int Index, int? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, ref int Index, decimal? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Format">保留格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, ref int Index, decimal? Value, string Format)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.Value.ToString(Format));
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, ref int Index, double? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Format">保留格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, ref int Index, double? Value, string Format)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.Value.ToString(Format));
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, ref int Index, float? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Format">保留格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, ref int Index, float? Value, string Format)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.Value.ToString(Format));
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, ref int Index, bool? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Format">时间格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, ref int Index, DateTime? Value, string Format = "yyyy-MM-dd")
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.Value.ToString(Format));
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Style">单元格样式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, ref int Index, string Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, int? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, decimal? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, decimal? Value, string Format, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value, Format);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, double? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, double? Value, string Format, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value, Format);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, float? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, float? Value, string Format, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value, Format);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, bool? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, DateTime? Value, ICellStyle Style, string Format = "yyyy-MM-dd")
        {
            var cell = Row.CreateTd(ref Index, Value, Format);
            cell.CellStyle = Style;
            return cell;
        }



        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列</param>
        /// <param name="Value">值</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, int Index, string Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (!string.IsNullOrEmpty(Value)) cell.SetCellValue(Value);
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, int Index, int? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, int Index, decimal? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Format">保留格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, int Index, decimal? Value, string Format)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.Value.ToString(Format));
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, int Index, double? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Format">保留格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, int Index, double? Value, string Format)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.Value.ToString(Format));
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, int Index, float? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Format">保留格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, int Index, float? Value, string Format)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.Value.ToString(Format));
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, int Index, bool? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Format">时间格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, int Index, DateTime? Value, string Format = "yyyy-MM-dd")
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.Value.ToString(Format));
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Style">单元格样式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, int Index, string Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, int? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, decimal? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, decimal? Value, string Format, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value, Format);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, double? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, double? Value, string Format, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value, Format);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, float? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, float? Value, string Format, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value, Format);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, bool? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, DateTime? Value, ICellStyle Style, string Format = "yyyy-MM-dd")
        {
            var cell = Row.CreateTd(ref Index, Value, Format);
            cell.CellStyle = Style;
            return cell;
        }

        
        #endregion


        /*

        #region 制作表格
               //表格制作
                var ExcelName="tableExport.xls";//excel文件名(含后缀)

                var book = new HSSFWorkbook();
                var sheet = book.CreateSheet("Sheet1");
                //表头样式
                ICellStyle HeadStyle = book.CreateCellStyle();
                HSSFFont font = (HSSFFont)book.CreateFont();
                font.FontName = "黑体";//字体
                font.Boldweight = 700;//加粗
                font.Color = HSSFColor.Black.Index;//颜色
                CellStyle.VerticalAlignment = VerticalAlignment.Center;//垂直居中 
                CellStyle.Alignment = HorizontalAlignment.CenterSelection;//水平居中
                HeadStyle.SetFont(font);

                 //单元格样式
                ICellStyle CellStyle = book.CreateCellStyle();
                CellStyle.VerticalAlignment = VerticalAlignment.Center;//垂直居中 
                CellStyle.Alignment = HorizontalAlignment.CenterSelection;//水平居中
                CellStyle.SetFont(book.CreateFont());

                //合并单元格
                //sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 12));

                //创建表头数组
                 var head = new string[] {
                    "利润中心", "县支公司代码", "对同一供应商的付款上限值", "年份", "方案号","类别",
                    "修改日期","修改人工号"
                };
                var rowIndex = 0; //行索引
                var hrow = sheet.CreateRow(rowIndex);//创建表头
                for (int i = 0; i < head.Length; i++)
                {
                    var columnIndex = i;
                    hrow.CreateTd(ref columnIndex, head[i],HeadStyle);
                }
                //循环数据列表
                foreach (var item in list)
                {
                    rowIndex++;
                    var trow = sheet.CreateRow(rowIndex); 
                    var columnIndex = 0;//列索引
                    trow.CreateTd(ref columnIndex, item.branch_no,CellStyle);//创建单元格
                }

                //自动列宽
                for (int i = 0; i < head.Length; i++)
                    sheet.AutoSizeColumn(i, true);
               //输出文件流
                using (var ms = new MemoryStream())
                {
                    book.Write(ms);
                    var value = ms.ToArray();

                    book.Close();
                    return File(value, "application/vnd.ms - excel",ExcelName);
                }
                 #endregion

        */

        #endregion
    }
}