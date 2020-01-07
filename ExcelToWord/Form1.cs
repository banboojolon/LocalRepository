using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using word = Microsoft.Office.Interop.Word;
using excel = Microsoft.Office.Interop.Excel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace ExcelToWord
{
    //qq号:956899912@qq.com  可能有验证信息，输入甜甜 应该是可以的   欢迎来找我
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 生成的文档  字体 比较小  请放大看   因为列数比较多  所以字体调小了
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button1_Click(object sender, EventArgs e)
        {
            //插入到word相应位置
            object sname = Environment.CurrentDirectory + "\\Test\\muban.docx";

            object oMissing = System.Reflection.Missing.Value;
            word.Application wordApp;
            wordApp = new word.Application();
            wordApp.Visible = false;//使文档可见
            word.Document doc = wordApp.Documents.Add(ref sname, ref oMissing, ref oMissing, ref oMissing);

            doc.Activate();

            doc.SpellingChecked = false;//关闭拼写检查
            doc.ShowSpellingErrors = false;//关闭显示拼写错误提示框 

            DateTime dt = DateTime.Now;

            object oBookmark = "Test";
            ISheet sheet;
            string filepathname = Environment.CurrentDirectory + "\\Test\\Test.xlsx";
            IWorkbook workbook = null;

            using (var fs = new FileStream(filepathname, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(fs);
            }
            sheet = workbook.GetSheetAt(0);

            int iFirstRowno = 0;//起始行
            int iFirstDataRowno = 5;//数据判断起始数据行  为了后面判断行删除做准备  这里假设判断删除的逻辑是：从数据部分开始  如果是空  0  -  都视为无效数据  如果一行的数据部分都是此规则  则删除此行  列同理
            int iLastColumn = 32;//列数

            //思路是先把数据放在dic里面  然后对dic进行处理  
            //当需要删除行列时，删除完之后需要对合并涉及到的行列进行  起止index重新计算  GetColumnMerge   GetRowMerge  这两个方法是核心算法  
            TestExcelToWord(doc, sheet, oBookmark.ToString(), workbook, iFirstRowno, 18, iFirstDataRowno, iLastColumn);
            object filename = Environment.CurrentDirectory + "\\Test\\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx";
            doc.SaveAs2(ref filename, ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);
            doc.Close(ref oMissing, ref oMissing, ref oMissing);
            GC.Collect();
        }

        private int TestExcelToWord(word.Document document, ISheet sheet, string bookmarks, IWorkbook workbook, int firstrowno, int lastrowno, int rowdatastart, int columnscount, bool isdelnull = true, int fixedcolumn = 4)
        {
            //表格行
            int iRowscount = lastrowno;
            //获取标签
            word.Bookmark bsbk = document.Bookmarks[bookmarks];
            word.Range bsrange = bsbk.Range;
            //表格列
            int iColumnscount = columnscount;

            //去掉上面的行数   数据开始的行号   数据结束的列号
            int iRowStart = firstrowno, iRowDataStart = rowdatastart, iColumDataStart = 2;

            //发生过异动的行号
            List<int> lstRowsno = new List<int>();
            //发生过异动的行号
            List<int> lstColumnsno = new List<int>();

            try
            {
                //循环sheet  然后后删除 这行是记录新旧编号的对应关系的
                Dictionary<int, int> dicOldNewColumn = new Dictionary<int, int>();
                //获取sheet原本信息
                var vdic = GetSheetInfo(sheet, iRowscount, iRowStart, iColumnscount, out Dictionary<int, List<rowcolumMerge>> dicmerge, dicOldNewColumn, false);

                Dictionary<int, List<string>> _lstmainvdic = new Dictionary<int, List<string>>(vdic);
                Dictionary<int, List<string>> _lsttempvdic = new Dictionary<int, List<string>>(vdic);

                //如何有合并  就删除空行列和调整行列编号
                if (dicmerge.Count > 0)
                {
                    vdic = GetNotAllRowNull(iRowDataStart, iColumDataStart, vdic, out lstRowsno);
                    vdic = GetNotAllColumnNull(iRowDataStart, iColumDataStart, vdic, out lstColumnsno, iRowStart, _lsttempvdic, out dicOldNewColumn, true);

                    if (lstRowsno.Count > 0)
                        //开始处理合并信息 调整行次列次
                        GetRowMerge(lstRowsno, dicmerge, iRowStart);
                    if (lstColumnsno.Count > 0)
                        GetColumnMerge(lstColumnsno, dicmerge);
                }

                //生成 doc 表格    
                word.Table tbl = bsrange.Tables.Add(bsrange, vdic.Count, vdic[iRowStart].Count);

                for (int columnno = 0; columnno < vdic[iRowStart].Count; columnno++)
                {
                    tbl.Columns[columnno + 1].Width = 30;
                }

                //表格的外边框显示
                tbl.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                ////表格的内边框显示
                tbl.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;

                //开始进行处理段落
                word.Paragraph tblpa;
                //单元格值
                var vCelltext = "";

                //计数  行次         实际的列
                int iHC = iRowStart, iRealColumncount = 0;

                int iCellRowno = 0, iCellColumno = 0;

                int iMergeColumn = 0;//合并列的信息行信息

                bool bFlag = true, bFlagRow = true;//标记是否需要继续列合并 行合并 

                Dictionary<int, List<ColumnSpan>> _cs = new Dictionary<int, List<ColumnSpan>>();
                List<ColumnSpan> columnSpans = new List<ColumnSpan>();
                NPOI.SS.UserModel.ICell iCell = null;
                bool fzFlag = false;


                /*将上年金额显示在表格下方：
                 * 1、找到上年金额所在的列 到此列开始   不在进行表格输入
                 * 2、修改合并的信息
                   3、组table 放在上年的表格下方  空一行   最后再把这行删掉*/

                foreach (int keys in vdic.Keys)
                {
                    iMergeColumn = 0;
                    // iLastColumnno = 0;
                    iRealColumncount = 0;

                    bFlagRow = true;
                    fzFlag = false;
                    columnSpans = new List<ColumnSpan>();
                    for (int columnno = 0; columnno < vdic[iRowStart].Count; columnno++)
                    {
                        ColumnSpan _objcolumnSpan = new ColumnSpan();
                        bFlag = true;
                        //初始情况或者是列大于这个合并的列
                        if ((iMergeColumn == 0 || columnno >= iMergeColumn) && iMergeColumn < vdic[keys].Count)
                        {
                            vCelltext = vdic[keys][columnno];

                            //单元格 字段 范围  书签
                            //如果系数大于1   由于标题重复原因  是从第四行开始读  所以需要rowno-1才是标签位置
                            iCellRowno = iHC - iRowStart + 1;
                            iCellColumno = iRealColumncount + 1;

                            word.Cell _wCell = tbl.Cell(iCellRowno, iCellColumno);
                            word.Paragraph cpara = _wCell.Range.Paragraphs.First;
                            word.Range range = _wCell.Range;
                            word.Bookmark bookm = range.Bookmarks.Add(bookmarks + "_" + iCellRowno + "_" + iCellColumno);

                            //对应excel中的单元格  如果设置格式  比如对齐方式  或者是否加粗 可通过icell获取
                            iCell = sheet.GetRow(keys).GetCell(dicOldNewColumn[columnno], MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            //获取对应excel中的单元格 样式
                            ICellStyle cellStyle = iCell.CellStyle;
                            IFont font = cellStyle.GetFont(workbook);

                            if (font.IsBold)
                            {
                                range.Font.Bold = font.Boldweight;
                            }
                            range.Font.Name = font.FontName;
                            range.Font.Size = 4;

                            word.Range _range = cpara.Range;
                            _range.Font.Name = font.FontName;
                            _range.Text = vCelltext;

                            int icsRowspan = 0;//初始状态
                                               //对单元格进行合并 列合并
                            if (dicmerge.ContainsKey(keys))//说明这行有需要合并的
                            {
                                dicmerge.Where(x => x.Key == keys).ToList()
                                    .ForEach(gg =>
                                    {
                                        if (bFlag)//如果这列没有合并过  如果这列已经合并  不需要走以下代码
                                        {
                                            dicmerge[gg.Key].ForEach(w =>
                                            {
                                                //合并列>1 需要合并 有列需要合并
                                                if (w.ColumnSpan > 1 && bFlag)//ocolumnindex
                                                {
                                                    if (columnno >= w.FirstColumnIndex && columnno <= w.LastColumnIndex)
                                                    {
                                                        try
                                                        {
                                                            //得到列信息 之后列循环   凡是坐落在此区间的   不再进行合并
                                                            iMergeColumn = columnno + w.ColumnSpan;
                                                            //iLastColumnno = columnno;
                                                            var ss = tbl.Cell(iCellRowno, iRealColumncount + w.ColumnSpan);
                                                            _wCell.Merge(ss);
                                                            //合并后居中
                                                            _wCell = tbl.Cell(iCellRowno, iCellColumno);
                                                            cpara = _wCell.Range.Paragraphs.First;

                                                            _wCell.VerticalAlignment = word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                                                            cpara.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                                                            _objcolumnSpan.Columnno = iCellColumno;
                                                            icsRowspan = w.RowSpan;
                                                            _objcolumnSpan.Rowspan = w.RowSpan;
                                                            _objcolumnSpan.FirstRowIndex = w.FirstRowIndex;
                                                            _objcolumnSpan.LastRowIndex = w.LastRowIndex;
                                                            _objcolumnSpan.Columnno = iCellColumno;
                                                            _objcolumnSpan.LastColumnIndex = w.LastColumnIndex;
                                                            _objcolumnSpan.FirstColumnIndex = w.FirstColumnIndex;
                                                            //单元格合并之后  需要把合并的单元格的列号进行更新  方便进行行合并
                                                            bFlag = false;
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            MessageBox.Show(ex.Message + "||" + ex.StackTrace.ToString(), "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                        }
                                                    }
                                                }

                                            });
                                        }
                                    });
                            }

                            if (icsRowspan == 0 && bFlag && dicmerge.ContainsKey(keys))//没有进行合并   找到对应的rowspan  用实际的columnno
                            {
                                bool isXH = true;

                                dicmerge[keys].ForEach(gg =>
                                {
                                    if (columnno >= gg.FirstColumnIndex && columnno <= gg.LastColumnIndex && isXH)
                                    {
                                        //在区间  
                                        _objcolumnSpan.Rowspan = gg.RowSpan;
                                        _objcolumnSpan.FirstRowIndex = gg.FirstRowIndex;
                                        _objcolumnSpan.LastRowIndex = gg.LastRowIndex;
                                        _objcolumnSpan.Columnno = iCellColumno;
                                        _objcolumnSpan.LastColumnIndex = gg.LastColumnIndex;
                                        _objcolumnSpan.FirstColumnIndex = gg.FirstColumnIndex;
                                        isXH = false;
                                    }
                                });
                            }

                            iRealColumncount++;
                            columnSpans.Add(_objcolumnSpan);
                        }
                    }
                    iHC++;
                    _cs[iCellRowno] = columnSpans;
                }

                columnSpans = new List<ColumnSpan>();
                Dictionary<string, string> _dictemp = new Dictionary<string, string>();

                //处理行合并
                foreach (int irowsno in _cs.Keys)
                {
                    columnSpans = _cs[irowsno];

                    if (columnSpans.Count > 0)
                    {
                        //循环合并
                        columnSpans.ForEach(x =>
                        {
                            if (x.Rowspan > 1 && x.LastColumnIndex < vdic[iRowStart].Count)
                            {
                                //判断是否已经合并过了
                                if (!_dictemp.ContainsKey(x.FirstRowIndex + "||" + x.LastRowIndex + "||" + x.LastColumnIndex + "||" + x.FirstColumnIndex))
                                {
                                    tbl.Cell(irowsno, x.Columnno).Merge(tbl.Cell(x.LastRowIndex - iRowStart + 1, x.LastColumnIndex + 1));//irowsno + x.Rowspan - 1
                                    _dictemp[x.FirstRowIndex + "||" + x.LastRowIndex + "||" + x.LastColumnIndex + "||" + x.FirstColumnIndex] = "";
                                }
                            }
                        });
                    }
                }

            }
            catch (Exception ex)
            {
                if (workbook != null)
                {
                    workbook.Close();
                }
                GC.Collect();
                MessageBox.Show(ex.Message + "||" + ex.StackTrace.ToString(), "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close();
                }
                GC.Collect();
            }

            return 1;
        }

        /// <summary>
        /// 剔除整行为空的行列  并处理合并的行列索引
        /// </summary>
        /// <param name="irowdatastart">数据开始的行数</param>
        /// <param name="icolumndatastart">数据体开始的列数</param>
        /// <param name="dicsheetinfo"></param>
        /// <param name="rowsno"></param>
        /// <returns></returns>
        private Dictionary<int, List<string>> GetNotAllRowNull(int irowdatastart, int icolumndatastart, Dictionary<int, List<string>> dicsheetinfo, out List<int> rowsno)
        {
            Dictionary<int, List<string>> _lstdicsheetinfo = new Dictionary<int, List<string>>(dicsheetinfo);// dicsheetinfo;
                                                                                                             //_lstdicsheetinfo = dicsheetinfo;
            rowsno = new List<int>();
            try
            {
                //临时使用的变量
                List<string> _lstTemp = new List<string>();
                //记录为空的个数
                int iEmptycount = 0;

                //剔除行信息  dic倒序循环  
                foreach (int idicrowno in dicsheetinfo.Keys)
                {
                    if (idicrowno >= irowdatastart)
                    {
                        _lstTemp = dicsheetinfo[idicrowno];
                        iEmptycount = 0;
                        for (int i = icolumndatastart; i < _lstTemp.Count; i++)
                        {
                            if (idicrowno == 39)
                            {
                            }
                            if (IsNull(_lstTemp[i]))
                            {
                                iEmptycount++;
                            }
                        }

                        //如果为空的数量=_lstTemp.Count-icolumndatastart  则说明都是空  此行可以删除
                        if (iEmptycount == _lstTemp.Count - icolumndatastart)
                        {
                            //删除 
                            _lstdicsheetinfo.Remove(idicrowno);
                            //记录删除的行号
                            rowsno.Add(idicrowno);

                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
            return _lstdicsheetinfo;
        }

        /// <summary>
        /// 剔除整列为空的行列  并处理合并的行列索引
        /// </summary>
        /// <param name="irowdatastart">数据开始的行数</param>
        /// <param name="icolumndatastart">数据体开始的列数</param>
        /// <param name="dicsheetinfo"></param>
        /// <param name="rowsno"></param>
        /// <returns></returns>
        private Dictionary<int, List<string>> GetNotAllColumnNull(int irowdatastart, int icolumndatastart, Dictionary<int, List<string>> dicsheetinfo, out List<int> columnsno, int irowstart, Dictionary<int, List<string>> _lsttempvdic, out Dictionary<int, int> dicOldNewColumn, bool isSyzqx = false)
        {
            Dictionary<int, List<string>> _lstdicsheetinfo = new Dictionary<int, List<string>>(dicsheetinfo);// dicsheetinfo; 
                                                                                                             //临时使用的变量
            List<string> _lstTemp = new List<string>();
            dicOldNewColumn = new Dictionary<int, int>();
            //记录为空的个数
            int iEmptycount = 0;
            columnsno = new List<int>();

            //剔除列信息 循环dic中的对应信息
            for (int icolumnno = icolumndatastart; icolumnno < _lsttempvdic[irowstart].Count; icolumnno++)
            {
                iEmptycount = 0;
                foreach (int keys in _lsttempvdic.Keys)
                {
                    if (keys >= irowdatastart)
                    {
                        if (IsNull(_lsttempvdic[keys][icolumnno]))
                        {
                            iEmptycount++;
                        }
                    }
                }

                //如果为空的数量=_lstTemp.Count-icolumndatastart  则说明都是空  此行可以删除
                if (iEmptycount == _lsttempvdic.Count - (irowdatastart - irowstart))
                {
                    //记录删除的列号
                    columnsno.Add(icolumnno);
                }

            }

            //所有者权益变动表  固定删除行次字段  
            if (isSyzqx) columnsno.Add(1);

            List<string> _lstTemp2 = new List<string>();
            //删除  
            foreach (int keys in dicsheetinfo.Keys)
            {
                _lstTemp = new List<string>();
                foreach (string strvalue in dicsheetinfo[keys])
                {
                    _lstTemp.Add(strvalue);
                }
                _lstTemp2 = new List<string>();
                for (int q = 0; q < _lstTemp.Count; q++)
                {
                    if (!columnsno.Contains(q))
                    {
                        dicOldNewColumn[_lstTemp2.Count] = q;//新列对应旧列的值
                        _lstTemp2.Add(_lstTemp[q]);
                    }
                }
                //每一个dic删除对应列的信息
                _lstdicsheetinfo[keys] = _lstTemp2;
            }
            return _lstdicsheetinfo;
        }

        /// <summary>
        /// 删除以后，行调整核心算法
        /// </summary>
        /// <param name="lstrowsno"></param>
        /// <param name="dicmerge"></param>
        /// <param name="irowstart"></param>
        private void GetRowMerge(List<int> lstrowsno, Dictionary<int, List<rowcolumMerge>> dicmerge, int irowstart)
        {
            /*根据删除的行 列  来计算合并需要调整的信息
                    * 并处理合并的最终结果
                    */

            for (int p = 0; p < lstrowsno.Count; p++)//循环删除的行
            {
                //如果删除的行是合并行
                if (dicmerge.ContainsKey(lstrowsno[p]))
                {
                    for (int dicp = irowstart; dicp < dicmerge.Count; dicp++)//循环合并的行
                    {
                        if (dicp > lstrowsno[p])//需要调整行
                        {
                            dicmerge[dicp].ForEach(x =>
                            {
                                //大于此行的所有  行索引减1
                                x.rowindex--;

                                //在区间里的数据处理
                                if (lstrowsno[p] >= x.oFirstRowIndex && lstrowsno[p] <= x.oLastRowIndex)
                                {
                                    //如果合并索引大于这行  则 索引直接--
                                    //if (lstrowsno[p] >= x.oFirstRowIndex)//等于或者大于都要--
                                    //{
                                    x.LastRowIndex--;
                                    x.RowSpan--;
                                    //}
                                }
                                else//不在区间里
                                {
                                    //只用处理行号大于删除行的  小于删除行的不变
                                    if (lstrowsno[p] < x.orowindex)
                                    {
                                        x.FirstRowIndex--;
                                        x.LastRowIndex--;
                                    }
                                }
                            });
                        }
                        else
                        //如果是小于此行的
                        {
                            dicmerge[dicp].ForEach(x =>
                            {
                                if (x.oLastRowIndex >= lstrowsno[p])
                                {
                                    x.LastRowIndex--;
                                    x.RowSpan--;
                                }
                            });
                        }
                    }
                }
                //如果删除的行不是合并行
                else
                {
                    for (int dicp = irowstart; dicp < dicmerge.Count; dicp++)//循环合并的行
                    {
                        if (dicp > lstrowsno[p])//需要调整行
                        {
                            dicmerge[dicp].ForEach(x =>
                            {
                                //大于此行的所有  行索引减1
                                x.rowindex--;
                                x.FirstRowIndex--;
                                x.LastRowIndex--;
                                x.RowSpan--;
                            });
                        }
                    }

                }
            }

        }

        /// <summary>
        /// 删除以后，列调整核心算法
        /// </summary>
        /// <param name="lstcolumnsno"></param>
        /// <param name="dicmerge"></param>
        private void GetColumnMerge(List<int> lstcolumnsno, Dictionary<int, List<rowcolumMerge>> dicmerge)
        {
            /*根据删除的行 列  来计算合并需要调整的信息
                    * 并处理合并的最终结果
                    */

            bool bMergeColumn = false;
            for (int p = 0; p < lstcolumnsno.Count; p++)//循环删除的行
            {
                foreach (int dicp in dicmerge.Keys)
                {
                    for (int imerge = 0; imerge < dicmerge[dicp].Count; imerge++)
                    {
                        if (dicmerge[dicp][imerge].ocolumnindex == lstcolumnsno[p])
                        {
                            bMergeColumn = true;
                            break;
                        }
                    }
                    //如果删除的行是合并行
                    if (bMergeColumn)
                    {
                        dicmerge[dicp].ForEach(g =>
                        {
                            var vss = dicmerge[dicp];
                            //在区间里的数据处理
                            if (lstcolumnsno[p] >= g.oFirstColumnIndex && lstcolumnsno[p] <= g.oLastColumnIndex)
                            {
                                g.LastColumnIndex--;
                                g.ColumnSpan--;

                            }
                            else//不在区间里
                            {
                                //只用处理行号大于删除行的  小于删除行的不变
                                if (lstcolumnsno[p] < g.ocolumnindex)
                                {
                                    g.FirstColumnIndex--;
                                    g.LastColumnIndex--;
                                }
                            }

                            if (g.ocolumnindex > lstcolumnsno[p])//需要调整行
                            {
                                //大于此行的所有  行索引减1
                                g.columnindex--;
                            }
                        });
                    }
                    //如果删除的行不是合并行
                    else
                    {
                        dicmerge[dicp].ForEach(q =>
                        {
                            if (q.ocolumnindex >= lstcolumnsno[p])//需要调整列
                            {
                                //大于此行的所有  行索引减1 
                                q.FirstColumnIndex--;
                                q.LastColumnIndex--;
                                q.columnindex--;

                                //横跨的列  判断是否需要减1
                                if (q.oFirstColumnIndex <= lstcolumnsno[p])
                                {
                                    q.ColumnSpan--;
                                }
                            }
                        });

                    }

                }//
            }
        }

        private Dictionary<int, List<string>> GetSheetInfo(ISheet sheet, int irowscount, int irowStart, int icolumnscount, out Dictionary<int, List<rowcolumMerge>> dicmerge, Dictionary<int, int> dicoddnewcolumn, bool isFZ)
        {
            //存放sheet的原始信息
            Dictionary<int, List<string>> dicSheetInfo = new Dictionary<int, List<string>>();
            //存放每行列信息
            List<string> _lstColumn = new List<string>();
            ICell cell;
            //存放  合并信息 
            dicmerge = new Dictionary<int, List<rowcolumMerge>>();
            List<rowcolumMerge> _lstrowcolumMerges = new List<rowcolumMerge>();

            bool bisMerge = false;

            //读取sheet数据放在dic中
            for (int isheetrowno = irowStart; isheetrowno < irowscount; isheetrowno++)
            {
                _lstColumn = new List<string>();
                _lstrowcolumMerges = new List<rowcolumMerge>();
                bisMerge = false;
                for (int isheetcolumno = 0; isheetcolumno < icolumnscount; isheetcolumno++)
                {
                    cell = sheet.GetRow(isheetrowno).GetCell(isheetcolumno, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    //如果是合并的行列 则记录下合并的信息
                    if (ExcelExtension.IsMergeCell(cell, out Dimension dimension))
                    {
                        //记录下行列合并的信息
                        _lstrowcolumMerges.Add(new rowcolumMerge
                        {
                            //word 里面以1开始作为索引开始
                            orowindex = isheetrowno,
                            ocolumnindex = isheetcolumno,

                            rowindex = isheetrowno,
                            columnindex = isheetcolumno,
                            DataCell = dimension.DataCell,
                            RowSpan = dimension.RowSpan,
                            ColumnSpan = dimension.ColumnSpan,

                            FirstRowIndex = dimension.FirstRowIndex,
                            LastRowIndex = dimension.LastRowIndex,
                            FirstColumnIndex = dimension.FirstColumnIndex,
                            LastColumnIndex = dimension.LastColumnIndex,


                            oFirstRowIndex = dimension.FirstRowIndex,
                            oLastRowIndex = dimension.LastRowIndex,
                            oFirstColumnIndex = dimension.FirstColumnIndex,
                            oLastColumnIndex = dimension.LastColumnIndex

                        });
                        bisMerge = true;
                    }
                    _lstColumn.Add(GetCellTextFZ(cell, isheetcolumno));

                    dicoddnewcolumn[isheetcolumno] = isheetcolumno;
                }
                dicSheetInfo.Add(isheetrowno, _lstColumn);
                if (bisMerge) dicmerge.Add(isheetrowno, _lstrowcolumMerges);
            }
            return dicSheetInfo;
        }

        /// <summary>
        /// 删除的规则  在此处
        /// </summary>
        /// <param name="cellvalue"></param>
        /// <returns></returns>
        private bool IsNull(string cellvalue)
        {
            if (string.IsNullOrWhiteSpace(cellvalue) || cellvalue.Equals("0") || cellvalue.Equals("-"))
            {
                return true;
            }
            return false;
        }

        public static string GetCellTextFZ(ICell cell, int isheetcolumno)
        {
            string strCelltext = "";
            switch (cell.CellType)
            {
                case CellType.String:
                    strCelltext = cell.StringCellValue;
                    break;
                case CellType.Numeric:
                case CellType.Unknown:
                case CellType.Formula:
                    if (isheetcolumno == 0)//第一列
                    {
                        strCelltext = cell.NumericCellValue.ToString();
                    }
                    else
                    {
                        strCelltext = cell.NumericCellValue > 0.00 ?
                                                        cell.NumericCellValue < 1 ? "0" + cell.NumericCellValue.ToString("###,###.00") : cell.NumericCellValue.ToString("###,###.00")
                                                 : cell.NumericCellValue == 0.00 ? "-" :
                                                        cell.NumericCellValue > -1 ? "(0" + (0 - cell.NumericCellValue).ToString("###,###.00") + ")" :
                                                 "(" + (0 - cell.NumericCellValue).ToString("###,###.00") + ")";
                    }
                    break;
                default:
                    strCelltext = "";
                    break;
                case CellType.Blank:
                    strCelltext = "";
                    break;
            }
            return strCelltext;
        } 
    }

    public class rowcolumMerge
    {
        /// <summary>
        /// 原始的行
        /// </summary>
        public int orowindex { get; set; }
        //原始的列
        public int ocolumnindex { get; set; }
        /// <summary>
        /// 最终的行
        /// </summary>
        public int rowindex { get; set; }
        //最终的列
        public int columnindex { get; set; }

        /// <summary>
        /// 含有数据的单元格(通常表示合并单元格的第一个跨度行第一个跨度列)，该字段可能为null
        /// </summary>
        public ICell DataCell { get; set; }

        /// <summary>
        /// 行跨度(跨越了多少行)
        /// </summary>
        public int RowSpan { get; set; }

        /// <summary>
        /// 列跨度(跨越了多少列)
        /// </summary>
        public int ColumnSpan { get; set; }

        /// <summary>
        /// 合并单元格的起始行索引
        /// </summary>
        public int FirstRowIndex { get; set; }

        /// <summary>
        /// 合并单元格的结束行索引
        /// </summary>
        public int LastRowIndex { get; set; }

        /// <summary>
        /// 合并单元格的起始列索引
        /// </summary>
        public int FirstColumnIndex { get; set; }

        /// <summary>
        /// 合并单元格的结束列索引
        /// </summary>
        public int LastColumnIndex { get; set; }


        /// <summary>
        /// 合并单元格的起始行索引
        /// </summary>
        public int oFirstRowIndex { get; set; }

        /// <summary>
        /// 合并单元格的结束行索引
        /// </summary>
        public int oLastRowIndex { get; set; }

        /// <summary>
        /// 合并单元格的起始列索引
        /// </summary>
        public int oFirstColumnIndex { get; set; }

        /// <summary>
        /// 合并单元格的结束列索引
        /// </summary>
        public int oLastColumnIndex { get; set; }
    }

    public class ColumnSpan
    {
        public int Columnno { get; set; }
        public int Rowspan { get; set; }
        /// <summary>
        /// 合并单元格的起始行索引
        /// </summary>
        public int FirstRowIndex { get; set; }

        /// <summary>
        /// 合并单元格的结束行索引
        /// </summary>
        public int LastRowIndex { get; set; }
        /// <summary>
        ///  
        /// </summary>
        public int LastColumnIndex { get; set; }
        public int FirstColumnIndex { get; set; }

        public string CellText { get; set; }
    }
}
