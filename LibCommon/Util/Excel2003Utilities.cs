using log4net;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace LibCommon.Util
{
    /// <summary>
    /// 使用DataTable或class直接匯出時會套用的style種類
    /// </summary>
    public enum CellStyleType
    {
        /// <summary>
        /// 表頭
        /// </summary>
        Header = 1,

        /// <summary>
        /// 奇數列
        /// </summary>
        Odd = 2,

        /// <summary>
        /// 偶數列
        /// </summary>
        Even = 3
    }

    /// <summary>
    /// 設定物件匯出的Attribute, 加在class的property上
    /// </summary>
    public class ExcelExportAttribute : Attribute
    {
        /// <summary>
        /// 欄位名稱
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 欄位排序, 排序時由小至大
        /// </summary>
        public int Order { get; set; }
    }

    public class Excel2003Utilities
    {
        private IWorkbook workBook;

        private IDictionary<int, int> sheetCurrentRowDict;

        private ILog logger = LogManager.GetLogger(typeof(Excel2003Utilities));

        /// <summary>
        /// 建立只有一個工作表的Excel
        /// </summary>
        public Excel2003Utilities()
        {
            workBook = new HSSFWorkbook();
            sheetCurrentRowDict = new Dictionary<int, int>();

            ISheet sheet = workBook.CreateSheet("Sheet1");
            sheet.ForceFormulaRecalculation = true;
            sheetCurrentRowDict.Add(0, 0);
        }

        /// <summary>
        /// 建立指定數量工作表的Excel
        /// </summary>
        /// <param name="sheetCount"></param>
        public Excel2003Utilities(int sheetCount)
        {
            workBook = new HSSFWorkbook();
            sheetCurrentRowDict = new Dictionary<int, int>();

            for (int i = 0; i <= sheetCount - 1; i++)
            {
                string sheetName = string.Format("Sheet{0}", i + 1);
                ISheet sheet = workBook.CreateSheet(sheetName);
                sheet.ForceFormulaRecalculation = true;
                sheetCurrentRowDict.Add(i, 0);
            }
        }

        /// <summary>
        /// 取得第一個工作表的目前列數
        /// </summary>
        /// <returns></returns>
        public int GetCurrentRow()
        {
            return GetCurrentRow(0);
        }

        /// <summary>
        /// 取得指定工作表的目前列數
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <returns></returns>
        public int GetCurrentRow(int sheetIndex)
        {
            int currentRow = 0;
            sheetCurrentRowDict.TryGetValue(sheetIndex, out currentRow);
            return currentRow;
        }

        /// <summary>
        /// 設定指定工作表的名稱
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="name">名稱</param>
        public void SetSheetName(int sheetIndex, string name)
        {
            workBook.SetSheetName(sheetIndex, name);
        }

        /// <summary>
        /// 設定第一個工作表為直印或橫印
        /// </summary>
        /// <param name="isLandScape">true為橫印, false為直印</param>
        public void SetLandScape(bool isLandScape)
        {
            SetLandScape(0, isLandScape);
        }

        /// <summary>
        /// 設定指定的工作表為直印或橫印
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="isLandScape">true為橫印, false為直印</param>
        public void SetLandScape(int sheetIndex, bool isLandScape)
        {
            workBook.GetSheetAt(sheetIndex).PrintSetup.Landscape = isLandScape;
        }

        /// <summary>
        /// 設定所有的工作表為直印或橫印
        /// </summary>
        /// <param name="isLandScape">true為橫印, false為直印</param>
        public void SetLandScapeToAllSheets(bool isLandScape)
        {
            for (int i = 0; i <= workBook.NumberOfSheets - 1; i++)
            {
                SetLandScape(i, isLandScape);
            }
        }

        /// <summary>
        /// 設定第一個工作表的指定區塊為指定的值
        /// </summary>
        /// <param name="value">要設定的值</param>
        /// <param name="firstRow">起始列</param>
        /// <param name="lastRow">結束列</param>
        /// <param name="firstColumn">起始欄</param>
        /// <param name="lastColumn">結束欄</param>
        public void SetValue(string val, int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            SetValue(0, val, firstRow, lastRow, firstColumn, lastColumn);
        }

        /// <summary>
        /// 設定指定的工作表的指定區塊為指定的值
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="value">要設定的值</param>
        /// <param name="firstRow">起始列</param>
        /// <param name="lastRow">結束列</param>
        /// <param name="firstColumn">起始欄</param>
        /// <param name="lastColumn">結束欄</param>
        public void SetValue(int sheetIndex, string val, int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            ISheet sheet = workBook.GetSheetAt(sheetIndex);

            for (int i = 0; firstRow + i <= lastRow; i++)
            {
                for (int j = 0; firstColumn + j <= lastColumn; j++)
                {
                    IRow row = sheet.GetRow(firstRow + i);
                    if (row == null)
                    {
                        row = sheet.CreateRow(firstRow + i);
                    }

                    ICell cell = row.GetCell(firstColumn + j);

                    if (cell == null)
                    {
                        cell = row.CreateCell(firstColumn + j);
                    }


                    cell.SetCellType(CellType.Blank);

                    double num = 0;
                    if (double.TryParse(val, out num) == true)
                    {
                        //判斷是否為自訂編碼, 以0開頭但不為小數
                        if (val.StartsWith("0") == true && val.StartsWith("0.") == false)
                        {
                            cell.SetCellValue(val);
                        }
                        else
                        {
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(num);
                        }
                    }
                    else
                    {
                        cell.SetCellValue(val);
                    }
                }
            }
        }

        /// <summary>
        /// 設定第一個工作表的指定區塊為指定的值, 使用IRichTextString, 可在同一儲存格內套用不同字型
        /// </summary>
        /// <param name="richString">要設定的IRichTextString</param>
        /// <param name="firstRow">起始列</param>
        /// <param name="lastRow">結束列</param>
        /// <param name="firstColumn">起始欄</param>
        /// <param name="lastColumn">結束欄</param>
        public void SetValue(IRichTextString richString, int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            SetValue(0, richString, firstRow, lastRow, firstColumn, lastColumn);
        }

        /// <summary>
        /// 設定指定的工作表的指定區塊為指定的值, 使用IRichTextString, 可在同一儲存格內套用不同字型
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="richString">要設定的IRichTextString</param>
        /// <param name="firstRow">起始列</param>
        /// <param name="lastRow">結束列</param>
        /// <param name="firstColumn">起始欄</param>
        /// <param name="lastColumn">結束欄</param>
        public void SetValue(int sheetIndex, IRichTextString richString, int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            ISheet sheet = workBook.GetSheetAt(sheetIndex);

            for (int i = 0; firstRow + i <= lastRow; i++)
            {
                for (int j = 0; firstColumn + j <= lastColumn; j++)
                {
                    IRow row = sheet.GetRow(firstRow + i);
                    if (row == null)
                    {
                        row = sheet.CreateRow(firstRow + i);
                    }

                    ICell cell = row.GetCell(firstColumn + j);

                    if (cell == null)
                    {
                        cell = row.CreateCell(firstColumn + j);
                    }

                    //RichTextString無法設定SetCellType
                    cell.SetCellValue(richString);
                }
            }
        }

        /// <summary>
        /// 設定第一個工作表的指定區塊套用統一style
        /// </summary>
        /// <param name="style">style</param>
        /// <param name="firstRow">起始列</param>
        /// <param name="lastRow">結束列</param>
        /// <param name="firstColumn">起始欄</param>
        /// <param name="lastColumn">結束欄</param>
        public void SetStyle(ICellStyle style, int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            SetStyle(0, style, firstRow, lastRow, firstColumn, lastColumn);
        }

        /// <summary>
        /// 設定指定的工作表的指定區塊套用統一style
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="style">style</param>
        /// <param name="firstRow">起始列</param>
        /// <param name="lastRow">結束列</param>
        /// <param name="firstColumn">起始欄</param>
        /// <param name="lastColumn">結束欄</param>
        public void SetStyle(int sheetIndex, ICellStyle style, int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            if (style != null)
            {
                ISheet sheet = workBook.GetSheetAt(sheetIndex);

                for (int i = 0; firstRow + i <= lastRow; i++)
                {
                    for (int j = 0; firstColumn + j <= lastColumn; j++)
                    {
                        IRow row = sheet.GetRow(firstRow + i);
                        if (row == null)
                        {
                            row = sheet.CreateRow(firstRow + i);
                        }

                        ICell cell = row.GetCell(firstColumn + j);

                        if (cell == null)
                        {
                            cell = row.CreateCell(firstColumn + j);
                        }

                        cell.CellStyle = style;
                    }
                }
            }
        }

        /// <summary>
        /// 設定第一個工作表的指定區塊套用統一字型設定
        /// </summary>
        /// <param name="font">font</param>
        /// <param name="firstRow">起始列</param>
        /// <param name="lastRow">結束列</param>
        /// <param name="firstColumn">起始欄</param>
        /// <param name="lastColumn">結束欄</param>
        public void SetFont(IFont font, int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            SetFont(0, font, firstRow, lastRow, firstColumn, lastColumn);
        }

        /// <summary>
        /// 設定指定的工作表的指定區塊套用統一字型設定
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="font">font</param>
        /// <param name="firstRow">起始列</param>
        /// <param name="lastRow">結束列</param>
        /// <param name="firstColumn">起始欄</param>
        /// <param name="lastColumn">結束欄</param>
        public void SetFont(int sheetIndex, IFont font, int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            if (font != null)
            {
                ISheet sheet = workBook.GetSheetAt(sheetIndex);

                for (int i = 0; firstRow + i <= lastRow; i++)
                {
                    for (int j = 0; firstColumn + j <= lastColumn; j++)
                    {
                        IRow row = sheet.GetRow(firstRow + i);
                        if (row == null)
                        {
                            row = sheet.CreateRow(firstRow + i);
                        }

                        ICell cell = row.GetCell(firstColumn + j);

                        if (cell == null)
                        {
                            cell = row.CreateCell(firstColumn + j);
                        }

                        ICellStyle newStyle = workBook.CreateCellStyle();
                        if (cell.CellStyle != null)
                        {
                            newStyle.CloneStyleFrom(cell.CellStyle);
                        }
                        newStyle.SetFont(font);
                        cell.CellStyle = newStyle;
                    }
                }
            }
        }

        /// <summary>
        /// 設定第一個工作表的指定區塊套用統一文字格式
        /// </summary>
        /// <param name="format">format, ex: 小數六位 0.000000</param>
        /// <param name="firstRow">起始列</param>
        /// <param name="lastRow">結束列</param>
        /// <param name="firstColumn">起始欄</param>
        /// <param name="lastColumn">結束欄</param>
        public void SetDataFormat(string format, int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            SetDataFormat(0, format, firstRow, lastRow, firstColumn, lastColumn);
        }

        /// <summary>
        /// 設定指定的工作表的指定區塊套用統一文字格式
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="format">format, ex: 小數六位 0.000000</param>
        /// <param name="firstRow">起始列</param>
        /// <param name="lastRow">結束列</param>
        /// <param name="firstColumn">起始欄</param>
        /// <param name="lastColumn">結束欄</param>
        public void SetDataFormat(int sheetIndex, string format, int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            if (format != null)
            {
                ISheet sheet = workBook.GetSheetAt(sheetIndex);

                for (int i = 0; firstRow + i <= lastRow; i++)
                {
                    for (int j = 0; firstColumn + j <= lastColumn; j++)
                    {
                        IRow row = sheet.GetRow(firstRow + i);
                        if (row == null)
                        {
                            row = sheet.CreateRow(firstRow + i);
                        }

                        ICell cell = row.GetCell(firstColumn + j);

                        if (cell == null)
                        {
                            cell = row.CreateCell(firstColumn + j);
                        }

                        ICellStyle newStyle = workBook.CreateCellStyle();
                        if (cell.CellStyle != null)
                        {
                            newStyle.CloneStyleFrom(cell.CellStyle);
                        }
                        IDataFormat dataFormat = workBook.CreateDataFormat();
                        newStyle.DataFormat = dataFormat.GetFormat(format);
                        cell.CellStyle = newStyle;
                    }
                }
            }
        }

        /// <summary>
        /// 設定第一個工作表的指定儲存格套用公式, 如: SUM(A1:C2)
        /// </summary>
        /// <param name="formula">公式字串, 如: SUM(A1:C2)</param>
        /// <param name="row">列index</param>
        /// <param name="column">欄index</param>
        public void SetFormula(string formula, int row, int column)
        {
            SetFormula(0, formula, row, column);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="formula">公式字串, 如: SUM(A1:C2)</param>
        /// <param name="rowIndex">列index</param>
        /// <param name="columnIndex">欄index</param>
        public void SetFormula(int sheetIndex, string formula, int rowIndex, int columnIndex)
        {
            if (formula != null)
            {
                ISheet sheet = workBook.GetSheetAt(sheetIndex);
                IRow row = sheet.GetRow(rowIndex);
                if (row == null)
                {
                    row = sheet.CreateRow(rowIndex);
                }

                ICell cell = row.GetCell(columnIndex);

                if (cell == null)
                {
                    cell = row.CreateCell(columnIndex);
                }

                cell.SetCellType(CellType.Formula);
                cell.SetCellFormula(formula);
            }
        }

        /// <summary>
        /// 設定第一個工作表中指定的欄自動調整欄寬，有針對中文字體無效的問題
        /// </summary>
        /// <param name="columnIndexs"></param>
        public void SetAutoSize(IEnumerable<int> columnIndexs)
        {
            SetAutoSize(0, columnIndexs);
        }

        /// <summary>
        /// 設定指定的工作表中指定的欄自動調整欄寬，有針對中文字體無效的問題
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="columnIndexs">欄index</param>
        public void SetAutoSize(int sheetIndex, IEnumerable<int> columnIndexs)
        {
            for (int i = 0; i <= columnIndexs.Count() - 1; i++)
            {
                workBook.GetSheetAt(sheetIndex).AutoSizeColumn(columnIndexs.ElementAt(i));
            }
        }

        /// <summary>
        /// ===效能不佳請慎用===
        /// 設定指定的工作表中指定的欄，依據每欄最長字串自動調整欄寬        
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="columnIndexs">欄index</param>
        /// <remarks></remarks>
        public void SetAutoSizeCustomize(int sheetIndex, IEnumerable<int> columnIndexs)
        {
            SetAutoSizeCustomize(sheetIndex, columnIndexs, false);
        }

        /// <summary>
        /// ===效能不佳請慎用===
        /// 設定指定的工作表中指定的欄，依據每欄最長字串自動調整欄寬
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="columnIndexs">欄index</param>
        /// <param name="useMergedCells">自動調整時是否將Merge Cells列入計算</param>
        /// <remarks></remarks>
        public void SetAutoSizeCustomize(int sheetIndex, IEnumerable<int> columnIndexs, bool useMergedCells)
        {
            ISheet sheet = workBook.GetSheetAt(sheetIndex);

            for (int i = 0; i <= columnIndexs.Count() - 1; i++)
            {
                sheet.AutoSizeColumn(columnIndexs.ElementAt(i), useMergedCells);

                int columnWidth = sheet.GetColumnWidth(columnIndexs.ElementAt(i)) / 256;
                for (int rowNum = 1; rowNum <= sheet.LastRowNum; rowNum++)
                {
                    IRow currentRow = default(IRow);

                    if (sheet.GetRow(rowNum) == null)
                    {
                        currentRow = sheet.CreateRow(rowNum);
                    }
                    else
                    {
                        currentRow = sheet.GetRow(rowNum);
                    }

                    if (currentRow.GetCell(columnIndexs.ElementAt(i)) != null)
                    {
                        ICell currentCell = currentRow.GetCell(columnIndexs.ElementAt(i));
                        int length = Encoding.Default.GetBytes(currentCell.ToString()).Length;
                        if (columnWidth < length)
                        {
                            columnWidth = length;
                        }
                    }
                }

                //加1解決誤差問題
                sheet.SetColumnWidth(columnIndexs.ElementAt(i), (columnWidth + 1) * 256);
            }
        }

        /// <summary>
        /// 設定第一個工作表中指定的欄之欄寬
        /// </summary>
        /// <param name="columnIndex">欄index</param>
        /// <param name="width">寬度</param>
        /// <remarks></remarks>
        public void SetColumnWidth(int columnIndex, int width)
        {
            SetColumnWidth(0, columnIndex, width);
        }

        /// <summary>
        /// 設定指定的工作表中指定的欄之欄寬
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="columnIndex">欄index</param>
        /// <param name="width">寬度</param>
        /// <remarks></remarks>
        public void SetColumnWidth(int sheetIndex, int columnIndex, int width)
        {
            workBook.GetSheetAt(sheetIndex).SetColumnWidth(columnIndex, width);
        }

        /// <summary>
        /// 於第一個工作表中繼續寫入一列空白列
        /// </summary>
        public void AppendEmptyLine()
        {
            AppendEmptyLineBySheet(0);
        }

        /// <summary>
        /// 於第一個工作表中繼續寫入指定數目的空白列
        /// </summary>
        /// <param name="count">要寫入的列數</param>
        public void AppendEmptyLine(int count)
        {
            AppendEmptyLineBySheet(0, count);
        }

        /// <summary>
        /// 於指定工作表中繼續寫入一列空白列
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        public void AppendEmptyLineBySheet(int sheetIndex)
        {
            int currentRow = sheetCurrentRowDict[sheetIndex];
            workBook.GetSheetAt(sheetIndex).CreateRow(currentRow);
            sheetCurrentRowDict[sheetIndex] = currentRow + 1;
        }

        /// <summary>
        /// 於指定工作表中繼續寫入指定數目的空白列
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="count">要寫入的列數</param>
        public void AppendEmptyLineBySheet(int sheetIndex, int count)
        {
            for (int i = 0; i <= count - 1; i++)
            {
                AppendEmptyLineBySheet(sheetIndex);
            }
        }

        /// <summary>
        /// 於第一個工作表中繼續寫入一列
        /// </summary>
        /// <param name="line">儲存格的集合</param>
        public void AppendNewLine(IEnumerable<string> line)
        {
            AppendNewLine(0, line, null);
        }

        /// <summary>
        /// 於第一個工作表中繼續寫入一列並指定統一style
        /// </summary>
        /// <param name="line">儲存格的集合</param>
        /// <param name="style">style設定</param>
        public void AppendNewLine(IEnumerable<string> line, ICellStyle style)
        {
            AppendNewLine(0, line, style);
        }

        /// <summary>
        /// 於第一個工作表中繼續寫入一列並指定個別style，styles個數需和line元素個數相同
        /// </summary>
        /// <param name="line">儲存格的集合</param>
        /// <param name="styles">style設定集合</param>
        public void AppendNewLine(IEnumerable<string> line, IEnumerable<ICellStyle> styles)
        {
            AppendNewLineWithStyles(0, line, styles);
        }

        /// <summary>
        /// 於指定工作表中繼續寫入一列
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="line">儲存格的集合</param>
        public void AppendNewLine(int sheetIndex, IEnumerable<string> line)
        {
            AppendNewLine(sheetIndex, line, null);
        }

        /// <summary>
        /// 於指定工作表中繼續寫入一列並指定統一style
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="line">儲存格的集合</param>
        /// <param name="style">style設定</param>
        public void AppendNewLine(int sheetIndex, IEnumerable<string> line, ICellStyle style)
        {
            int currentRow = sheetCurrentRowDict[sheetIndex];
            IRow row = (HSSFRow)workBook.GetSheetAt(sheetIndex).CreateRow(currentRow);
            for (int i = 0; i <= line.Count() - 1; i++)
            {
                ICell cell = (HSSFCell)row.CreateCell(i);
                string val = line.ElementAt(i);

                double num = 0;
                if (double.TryParse(val, out num) == true)
                {
                    //判斷是否為自訂編碼, 以0開頭但不為小數
                    if (val.StartsWith("0") == true && val.StartsWith("0.") == false)
                    {
                        cell.SetCellValue(val);
                    }
                    else
                    {
                        cell.SetCellType(CellType.Numeric);
                        cell.SetCellValue(num);
                    }
                }
                else
                {
                    cell.SetCellValue(val);
                }

                if (style != null)
                {
                    cell.CellStyle = style;
                }
            }
            sheetCurrentRowDict[sheetIndex] = currentRow + 1;
        }

        /// <summary>
        /// 於指定工作表中繼續寫入一列並指定style，styles個數需和line元素個數相同
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="line">儲存格的集合</param>
        /// <param name="styles">style設定集合</param>
        public void AppendNewLineWithStyles(int sheetIndex, IEnumerable<string> line, IEnumerable<ICellStyle> styles)
        {
            int currentRow = sheetCurrentRowDict[sheetIndex];
            IRow row = (HSSFRow)workBook.GetSheetAt(sheetIndex).CreateRow(currentRow);
            for (int i = 0; i <= line.Count() - 1; i++)
            {
                ICell cell = (HSSFCell)row.CreateCell(i);
                string val = line.ElementAt(i);

                double num = 0;
                if (double.TryParse(val, out num) == true)
                {
                    //判斷是否為自訂編碼, 以0開頭但不為小數
                    if (val.StartsWith("0") == true && val.StartsWith("0.") == false)
                    {
                        cell.SetCellValue(val);
                    }
                    else
                    {
                        cell.SetCellType(CellType.Numeric);
                        cell.SetCellValue(num);
                    }
                }
                else
                {
                    cell.SetCellValue(val);
                }

                if (styles != null)
                {
                    if (i < styles.Count())
                    {
                        cell.CellStyle = styles.ElementAt(i);
                    }
                }
            }
            sheetCurrentRowDict[sheetIndex] = currentRow + 1;
        }

        /// <summary>
        /// 建立儲存格格式物件
        /// </summary>
        /// <returns></returns>
        public ICellStyle CreateCellStyle()
        {
            return (HSSFCellStyle)workBook.CreateCellStyle();
        }

        /// <summary>
        /// 建立字型物件，需配合使用HSSFCellStyle.SetFont()，設定至HSSFCellStyle中才會生效
        /// </summary>
        /// <returns></returns>
        public IFont CreateFont()
        {
            return (HSSFFont)workBook.CreateFont();
        }

        /// <summary>
        /// 建立文字格式物件，需配合使用HSSFCellStyle.DataFormat屬性，設定至HSSFCellStyle中才會生效
        /// </summary>
        /// <returns></returns>
        public IDataFormat CreateDataFormat()
        {
            return (HSSFDataFormat)workBook.CreateDataFormat();
        }

        /// <summary>
        /// 將第一個工作表中的儲存格區塊合併
        /// </summary>
        /// <param name="firstRow">起始列</param>
        /// <param name="lastRow">結束列</param>
        /// <param name="firstColumn">起始欄</param>
        /// <param name="lastColumn">結束欄</param>
        public void MergeRegion(int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            MergeRegion(0, firstRow, lastRow, firstColumn, lastColumn);
        }

        /// <summary>
        /// 將指定工作表中的儲存格區塊合併
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="firstRow">起始列</param>
        /// <param name="lastRow">結束列</param>
        /// <param name="firstColumn">起始欄</param>
        /// <param name="lastColumn">結束欄</param>
        public void MergeRegion(int sheetIndex, int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            workBook.GetSheetAt(sheetIndex).AddMergedRegion(new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn));
        }

        /// <summary>
        /// 以標題及DataTable的ColumnName及內容寫入第一個工作表的資料列
        /// 使用既定的style顯示
        /// </summary>
        /// <param name="title">標題</param>
        /// <param name="dt">資料DataTable</param>
        /// <param name="isDefaultStyle">是否使用預設格式</param>
        public void AppendDataTable(string title, DataTable dt, bool isDefaultStyle = true)
        {
            AppendDataTable(0, title, dt, isDefaultStyle);
        }

        /// <summary>
        /// 以標題及DataTable的ColumnName及內容寫入指定工作表的資料列
        /// 使用既定的style顯示
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="title">標題</param>
        /// <param name="dt">資料DataTable</param>
        /// <param name="isDefaultStyle">是否使用預設格式</param>
        public void AppendDataTable(int sheetIndex, string title, DataTable dt, bool isDefaultStyle = true)
        {
            IList<string> headers = new List<string>();
            for (int i = 0; i <= dt.Columns.Count - 1; i++)
            {
                headers.Add(dt.Columns[i].ColumnName);
            }

            AppendDataTableWithHeader(sheetIndex, title, headers, dt, isDefaultStyle);
        }

        /// <summary>
        /// 以標題, 自訂的表頭及DataTable的內容寫入第一個工作表的資料列
        /// 使用既定的style顯示
        /// </summary>
        /// <param name="title">標題</param>
        /// <param name="headers">標頭列儲存格集合</param>
        /// <param name="dt">資料DataTable</param>
        /// <param name="isDefaultStyle">是否使用預設格式</param>
        public void AppendDataTableWithHeader(string title, IEnumerable<string> headers, DataTable dt, bool isDefaultStyle = true)
        {
            AppendDataTableWithHeader(0, title, headers, dt, isDefaultStyle);
        }

        /// <summary>
        /// 以標題, 自訂的表頭及DataTable的內容寫入指定工作表的資料列
        /// 使用既定的style顯示
        /// </summary>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="title">標題</param>
        /// <param name="headers">標頭列儲存格集合</param>
        /// <param name="dt">資料DataTable</param>
        /// <param name="isDefaultStyle">是否使用預設格式</param>
        public void AppendDataTableWithHeader(int sheetIndex, string title, IEnumerable<string> headers, DataTable dt, bool isDefaultStyle = true)
        {
            ICellStyle headerStyle = GetCellStyle(CellStyleType.Header);
            ICellStyle oddStyle = GetCellStyle(CellStyleType.Odd);
            ICellStyle evenStyle = GetCellStyle(CellStyleType.Even);

            AppendNewLine(sheetIndex, new List<string> { title });

            if (isDefaultStyle == true)
                AppendNewLine(sheetIndex, headers, headerStyle);
            else
                AppendNewLine(sheetIndex, headers);

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                IList<string> line = new List<string>();
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {
                    line.Add(dt.Rows[i][j].ToString());
                }

                if (isDefaultStyle == true)
                    AppendNewLine(sheetIndex, line, i % 2 == 0 ? evenStyle : oddStyle);
                else
                    AppendNewLine(sheetIndex, line);
            }
        }

        /// <summary>
        /// 以標題及物件集合寫入第一個工作表的資料列
        /// </summary>
        /// <typeparam name="T">物件型別</typeparam>
        /// <param name="title">標題</param>
        /// <param name="datas">物件集合</param>
        /// <param name="isDefaultStyle">是否使用預設格式</param>
        public void AppendDataCollection<T>(string title, IEnumerable<T> datas, bool isDefaultStyle = true)
        {
            AppendDataCollection(0, title, datas, isDefaultStyle);
        }

        /// <summary>
        /// 以標題及物件集合寫入指定工作表的資料列
        /// </summary>
        /// <typeparam name="T">物件型別</typeparam>
        /// <param name="sheetIndex">工作表index</param>
        /// <param name="title">標題</param>
        /// <param name="datas">物件集合</param>
        /// <param name="isDefaultStyle">是否使用預設格式</param>
        public void AppendDataCollection<T>(int sheetIndex, string title, IEnumerable<T> datas, bool isDefaultStyle = true)
        {
            ICellStyle headerStyle = GetCellStyle(CellStyleType.Header);
            ICellStyle oddStyle = GetCellStyle(CellStyleType.Odd);
            ICellStyle evenStyle = GetCellStyle(CellStyleType.Even);

            var headerList = GetHeaderLine<T>();

            AppendNewLine(sheetIndex, new List<string> { title });

            if (isDefaultStyle == true)
                AppendNewLine(sheetIndex, headerList, headerStyle);
            else
                AppendNewLine(sheetIndex, headerList);

            for (int i = 0; i <= datas.Count() - 1; i++)
            {
                IList<string> line = GetOutputLine<T>(datas.ElementAt(i)).ToList();

                if (isDefaultStyle == true)
                    AppendNewLine(sheetIndex, line, i % 2 == 0 ? evenStyle : oddStyle);
                else
                    AppendNewLine(sheetIndex, line);
            }
        }

        /// <summary>
        /// 將excel輸出至Stream, 僅能輸出為.xls檔案
        /// </summary>
        /// <param name="stream"></param>
        public void WriteStream(Stream stream)
        {
            try
            {
                workBook.Write(stream);
            }
            catch (Exception ex)
            {
                logger.Fatal(ex.Message, ex);
                throw;
            }
        }

        /// <summary>
        /// 將Excel轉換成Table(第一列為Table欄位名稱)
        /// </summary>
        /// <param name="filePath">檔案路徑</param>
        /// <param name="isFirstRowHeader">第一列資料是否為標頭</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public DataTable LoadExcelToDT(string filePath, bool isFirstRowHeader = true)
        {
            DataTable dt = new DataTable();

            if (File.Exists(filePath))
            {
                try
                {
                    IWorkbook book = null;

                    //開啟檔案
                    using (FileStream fs = new FileStream(filePath, FileMode.Open))
                    {
                        if (filePath.EndsWith(".xls"))
                        {
                            book = new HSSFWorkbook(fs);
                        }
                        else if (filePath.EndsWith(".xlsx"))
                        {
                            book = new XSSFWorkbook(fs);
                        }


                        ISheet sheet = book.GetSheetAt(0);

                        IRow hRow = sheet.GetRow(0);
                        if (isFirstRowHeader == true)
                        {
                            //取得excel首列並作為dt的欄位名稱
                            for (int i = hRow.FirstCellNum; i <= hRow.LastCellNum - 1; i++)
                            {
                                dt.Columns.Add(hRow.GetCell(i).StringCellValue, typeof(string));
                            }
                        }
                        else
                        {
                            //使用行數作為dt的欄位名稱
                            for (int i = hRow.FirstCellNum; i <= hRow.LastCellNum - 1; i++)
                            {
                                dt.Columns.Add(i.ToString(), typeof(string));
                            }
                        }

                        //判斷是否需跳過第一列
                        int skipRowCount = 0;
                        if (isFirstRowHeader == true)
                            skipRowCount = 1;

                        //取得資料
                        for (int i = sheet.FirstRowNum + skipRowCount; i <= sheet.LastRowNum; i++)
                        {
                            DataRow row = dt.NewRow();

                            for (int j = sheet.GetRow(i).FirstCellNum; j <= sheet.GetRow(i).LastCellNum - 1; j++)
                            {
                                row[j] = sheet.GetRow(i).GetCell(j);
                            }

                            dt.Rows.Add(row);
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Fatal(ex.Message, ex);
                    throw;
                }
            }

            return dt;
        }

        /// <summary>
        /// 儲存每個型別之ExcelExportAttribute及Property的對應設定
        /// </summary>
        private IDictionary<Type, IDictionary<ExcelExportAttribute, PropertyInfo>> typeExportSettingDict =
                                        new Dictionary<Type, IDictionary<ExcelExportAttribute, PropertyInfo>>();

        /// <summary>
        /// 依照物件型別, 取得ExcelExportAttribute及Property的對應設定, 並儲存為Dictionary的格式後回傳
        /// </summary>
        /// <typeparam name="T">資料物件型別</typeparam>
        /// <returns>ExcelExportAttribute及Property的對應設定Dictionary</returns>
        private IDictionary<ExcelExportAttribute, PropertyInfo> GetExcelExportAttribute<T>()
        {
            Type type = typeof(T);

            //若目前記憶體中尚無設定, 重新使用反射取得設定後存入記憶體中
            if (typeExportSettingDict.ContainsKey(type) == false)
            {
                PropertyInfo[] memberInfos = typeof(T).GetProperties();

                //宣告儲存每一個ExcelExportAttribute及Property的對應設定
                IDictionary<ExcelExportAttribute, PropertyInfo> dict = new Dictionary<ExcelExportAttribute, PropertyInfo>();

                foreach (PropertyInfo info in memberInfos)
                {
                    ExcelExportAttribute attribute = info.GetCustomAttributes(typeof(ExcelExportAttribute), false).FirstOrDefault() as ExcelExportAttribute;

                    //有設定ExcelExportAttribute才存入dict中
                    if (attribute != null)
                    {
                        if (dict.ContainsKey(attribute) == false)
                        {
                            dict.Add(attribute, info);
                        }
                    }
                }

                typeExportSettingDict.Add(type, dict);
            }

            return typeExportSettingDict[type];
        }

        /// <summary>
        /// 依照ExcelExportAttribute的設定順序, 取得要輸出的物件標頭的字串集合
        /// </summary>
        /// <typeparam name="T">資料物件型別</typeparam>
        /// <returns>物件標頭的字串集合</returns>
        private IEnumerable<string> GetHeaderLine<T>()
        {
            IEnumerable<ExcelExportAttribute> excelExports = GetExcelExportAttribute<T>().Keys.OrderBy(p => p.Order);

            List<string> headerLine = excelExports.Select(p => p.Name).ToList();

            return headerLine;
        }

        /// <summary>
        /// 依照ExcelExportAttribute的設定順序, 取得要輸出的物件屬性值的字串集合
        /// </summary>
        /// <typeparam name="T">資料物件型別</typeparam>
        /// <param name="data">資料物件</param>
        /// <returns>物件屬性值的字串集合</returns>
        private IEnumerable<string> GetOutputLine<T>(T data)
        {
            IDictionary<ExcelExportAttribute, PropertyInfo> dict = GetExcelExportAttribute<T>();

            IEnumerable<ExcelExportAttribute> excelExports = dict.Keys.OrderBy(p => p.Order);

            List<string> outputLine = new List<string>();
            foreach (ExcelExportAttribute export in excelExports)
            {
                PropertyInfo info = dict[export];

                object objVal = info.GetValue(data, null);
                string val = objVal != null ? objVal.ToString() : "";
                outputLine.Add(val);
            }

            return outputLine;
        }

        /// <summary>
        /// 取得儲存格style
        /// </summary>
        /// <param name="type">類型</param>
        /// <returns></returns>
        private ICellStyle GetCellStyle(CellStyleType type)
        {
            ICellStyle style = (HSSFCellStyle)workBook.CreateCellStyle();

            //設定儲存格框線
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BottomBorderColor = HSSFColor.Black.Index;
            style.LeftBorderColor = HSSFColor.Black.Index;
            style.RightBorderColor = HSSFColor.Black.Index;
            style.TopBorderColor = HSSFColor.Black.Index;

            //設定文字對齊及自動換行
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            style.WrapText = true;

            //背景
            switch (type)
            {
                case CellStyleType.Header:
                    style.FillForegroundColor = HSSFColor.Indigo.Index;
                    IFont font = (HSSFFont)workBook.CreateFont();
                    font.Boldweight = Convert.ToInt16(FontBoldWeight.Bold);
                    font.Color = HSSFColor.White.Index;
                    style.SetFont(font);
                    break;

                case CellStyleType.Odd:
                    style.FillForegroundColor = HSSFColor.LightCornflowerBlue.Index;
                    break;

                case CellStyleType.Even:
                    style.FillForegroundColor = HSSFColor.White.Index;
                    break;

            }
            style.FillPattern = FillPattern.SolidForeground;

            return style;
        }
    }
}
