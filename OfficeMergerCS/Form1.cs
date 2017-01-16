using System;
using System.IO;
using System.Drawing;
using System.Collections;
using Point = System.Drawing.Point;
using System.Windows.Forms;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace OfficeMergerCS
{
    public partial class Form1 : Form
    {
        //引用Excel Application類別
        _Application myExcel = null;
        //引用活頁簿類別 
        _Workbook myBook = null;
        //引用工作表類別
        _Worksheet mySheet = null;
        //引用Range類別 
        Range myRange = null;
        //新的活頁簿類別
        _Workbook newBook = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (myArray == null)
            {
                MessageBox.Show("请先读取数据");
                return;
            }

            //開啟一個新的應用程式
            myExcel = new Excel.Application();
            //加入新的活頁簿
            myExcel.Workbooks.Add(true);
            //停用警告訊息
            myExcel.DisplayAlerts = true;
            //讓Excel文件可見 
            myExcel.Visible = true;
            //引用第一個活頁簿
            myBook = myExcel.Workbooks[1];
            //設定活頁簿焦點
            myBook.Activate();
            //加入新的工作表在第1張工作表之後 
            myBook.Sheets.Add(Type.Missing, myBook.Worksheets[1], 1, Type.Missing);
            //引用第一個工作表
            mySheet = (Worksheet)myBook.Worksheets[1];
            //命名工作表的名稱為 "Array"
            mySheet.Name = "Array";
            //設工作表焦點
            mySheet.Activate();
            int UpBound1 = myArray.GetUpperBound(0);//二維陣列數上限
            int UpBound2 = myArray.GetUpperBound(1);//二維陣列數上限
            //寫入報表名稱 
            myExcel.Cells[1, 4] = "全自动生成報表";
            //設定範圍 
            myRange = (Range)mySheet.Range[mySheet.Cells[2, 1], mySheet.Cells[UpBound1 + 2, UpBound2 + 1]];
            myRange.Select();
            //用陣列一次寫入資料 
            myRange.Value2 = myArray;
            //設定儲存路徑 
            string PathFile = Directory.GetCurrentDirectory() + @"\我的报表.xlsx";
            //另存活頁簿 
            myBook.SaveAs(PathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
                , XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //關閉活頁簿 
            //myBook.Close(false, Type.Missing, Type.Missing);
            ////關閉Excel 
            //myExcel.Quit();
            ////釋放Excel資源 
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(myExcel);
            myBook = null;
            mySheet = null;
            myRange = null;
            myExcel = null;

            GC.Collect();
        }
        private Object[,] myArray;

        private void btnOpen_Click(object sender, EventArgs e)
        {
            Stream mystream;
            OpenFileDialog openfiledialog1 = new OpenFileDialog();
            openfiledialog1.Multiselect = true;//允许同时选择多个文件
            //openfiledialog1.InitialDirectory = "c:\\";
            openfiledialog1.Filter = "All files(*.*)|*.*|Word2003(*.doc)|(*.doc)|2003xls files(*.xls)|*.xls|2007xlsx files(*.xlsx)|*.xlsx";
            openfiledialog1.FilterIndex = 1;
            openfiledialog1.RestoreDirectory = true;
            if (openfiledialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((mystream = openfiledialog1.OpenFile()) != null)
                    {
                        for (int fi = 0; fi < openfiledialog1.FileNames.Length; fi++)
                        {
                            lvFile.Items.Add(ExtractFileName(openfiledialog1.FileNames[fi]), 0);
                            lvFile.Items[fi].SubItems.Add(ExtractFileExt(openfiledialog1.FileNames[fi]) + "文件");
                            lvFile.Items[fi].SubItems.Add(openfiledialog1.FileNames[fi]);
                            lvFile.Items[fi].ImageIndex = 0;
                        }
                        mystream.Close();
                    }
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message);
                }
            }
        }
        //获得文件后缀
        private string ExtractFileExt(string fileName)
        {
            string strEName = fileName.Substring(fileName.LastIndexOf(".") + 1, (fileName.Length - fileName.LastIndexOf(".") - 1));
            return strEName;
        }
        //获得文件名
        private string ExtractFileName(string fileName)
        {
            string strName = fileName.Substring(fileName.LastIndexOf("\\") + 1, (fileName.LastIndexOf(".") -
                 fileName.LastIndexOf("\\") - 1));
            return strName;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            // 添加列// Set to details 
            lvFile.View = View.Tile;//.LargeIcon;

            lvFile.Columns.Add("001", 100, HorizontalAlignment.Left);
            lvFile.Columns.Add("002", 100, HorizontalAlignment.Left);
            lvFile.Columns.Add("003", 100, HorizontalAlignment.Left);

            lvFile.HeaderStyle = ColumnHeaderStyle.Nonclickable;
            // 显示大图标列表（小图标和这个差不多）    首先拽一个imagelist控件到Form中来，然后为这个控件添加图片
            lvFile.TileSize = new Size(150, 80); ;
            lvFile.LargeImageList = imageList1;

            //Details模式下，自动适应宽度,-1根据内容设置宽度,-2根据标题设置宽度.
            lvFile.Columns[0].Width = -2;
            lvFile.Columns[1].Width = -1;
            //禁止ListView中进行多项选中（禁用多选）
            lvFile.MultiSelect = false;
            //读取用户数据
            foreach (string lbItem in omSet.Default.ListBoxSetting)
            {
                lbContent.Items.Add(lbItem);
            }
            tbMainRange.Text = omSet.Default.MainRangeSetting;
            tbMainStart.Text = omSet.Default.MainRangeStartSetting;
            tbMainEnd.Text = omSet.Default.MainRangeEndSetting;
            tbDataTag.Text = omSet.Default.MainRangeDataTag;

        }
        // 为ListView设置鼠标右键选中事件。经常需要在右键选中某项时弹出浮动菜单用到。    首先为ListView控件添加MouseClick的Event，然后下面代码：
        private void listView1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {        //给ListView在鼠标右键选中的情况下添加浮动菜单：
                String str = lvFile.SelectedItems[0].Text;
                Point p = new Point(e.X, e.Y);
                contextMenuStrip1.Show(lvFile, p);
            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (tbAdd.Text.ToString().Trim() == "")
            {
                MessageBox.Show("it's empty!");
            }
            else
            {
                lbContent.Items.Add(tbAdd.Text.ToString());
                omSet.Default.ListBoxSetting.Add(tbAdd.Text.ToString());
            }
            tbAdd.Clear();
            tbAdd.Select();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            omSet.Default.MainRangeSetting = tbMainRange.Text;
            omSet.Default.MainRangeStartSetting = tbMainStart.Text;
            omSet.Default.MainRangeEndSetting = tbMainEnd.Text;
            omSet.Default.MainRangeDataTag = tbDataTag.Text;
            omSet.Default.Save();
            MessageBox.Show("保存成功");
        }
        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            omSet.Default.ListBoxSetting.Remove(lbContent.SelectedItem.ToString());
            this.lbContent.Items.Remove(lbContent.SelectedItem);
        }

        private void listBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            tbAdd.Select();
        }

        private void tbAdd_TextChanged(object sender, EventArgs e)
        {

        }

        //读取数据
        private void btRead_Click(object sender, EventArgs e)
        {
            int MAXLINE = 5000;
            int i = 0, j = 0, k = 0, m = 0;//m为总行数
            int fileCount = lvFile.Items.Count;
            string DataTag;
            int eCount = 0;//有效工作簿数
            int sCount = 0;//当前表中工作簿数
            Point point;
            Object missing = Type.Missing;

            int iCount = lbContent.Items.Count;
            //重点区域，范围型读取单元格区域
            RangeSelector mainRange = new RangeSelector(tbMainRange.Text);
            //预判断块读取还是固定位置读取，初始化总数组大小
            if (mainRange.getWidth() > 0)
                myArray = new String[MAXLINE, mainRange.getWidth() + iCount + 1];//最多千行
            else
                myArray = new String[MAXLINE, iCount + 1];//最多千行

            //開啟一個新的應用程式
            myExcel = new Excel.Application();
            for (i = 0; i < fileCount; i++)
            {
                //停用警告訊息
                myExcel.DisplayAlerts = false;
                //讓Excel文件可見
                myExcel.Visible = true;
                //引用第一個活頁簿
                myBook = myExcel.Workbooks.Open(lvFile.Items[i].SubItems[2].Text, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                //設定活頁簿焦點
                myBook.Activate();
                //判断所有工作簿
                sCount = myBook.Worksheets.Count;
                for (k = 1; k <= sCount; k++)
                {
                    //大表判断条件
                    if (cbSheetSelect.Text != "全部" && Int16.Parse(cbSheetSelect.Text) != k) continue;
                    //选择当前表
                    mySheet = (Worksheet)myBook.Worksheets[k];
                    //設工作表焦點
                    mySheet.Activate();
                    //特征值判断
                    if (tbSheetPos.Text != "")
                    {
                        point = pointPos(tbSheetPos.Text);
                        if (mySheet.Cells[point.Y, point.X].Value != tbSheetCont.Text) continue;
                    }
                    eCount++;
                    //备注列判断
                    if (tbDataTag.Text == "") DataTag = lvFile.Items[i].SubItems[0].Text;    //未设置备注默认使用文件名
                    else
                    {
                        Point tagpos = pointPos(tbDataTag.Text);
                        DataTag = Convert.ToString(mySheet.Cells[tagpos.Y, tagpos.X].Value);
                    }
                    string mainStart = tbMainStart.Text;
                    string mainEnd = tbMainEnd.Text;
                    //判断选择哪种模式
                    if (mainRange.Count() > 1)
                    {
                        mainRange = new RangeSelector(tbMainRange.Text);//重新恢复原区域值
                        //重点区域起始位置判断
                        Point nowPos = mainRange.getCurPos();
                        for (j = 0; j < mainRange.Count(); j++)
                        {
                            string myCell = Convert.ToString(mySheet.Cells[nowPos.Y, nowPos.X].Value);
                            if (mainStart == "") break;
                            if (myCell == mainStart) break;
                            mainRange.acc();
                        }
                        //mainRange.lineacc();    //移到关键字下一行
                        mainRange.SetStartVal(mainRange.getCurPos());
                        //读取内容
                        while (m < MAXLINE)    //最大读取行数上限估计
                        {
                            nowPos = mainRange.getCurPos();
                            string lineFirstCell = Convert.ToString(mySheet.Cells[nowPos.Y, nowPos.X].Value);
                            if (lineFirstCell == null|| lineFirstCell=="") break;   //首字为空
                            if (lineFirstCell == tbMainEnd.Text) break; //符合结束字符串
                            if (mainRange.pos > mainRange.Count()) break;//读取完了就退出
                            for (j = 0; j < mainRange.getWidth(); j++)//读取一行
                            {
                                point = mainRange.getCurPos();
                                myArray[m, j] = Convert.ToString(mySheet.Cells[point.Y, point.X].Value);    //不管什么类型都转为字符串
                                mainRange.acc();
                            }
                            myArray[m, j] = DataTag;
                            m++;
                        }
                    }
                    else
                    {
                        //准备读取单元格相关信息，固定位置读取单元格
                        if (iCount >= 1)
                        {
                            List<Array> ListOfLine = new List<Array>(); //所有的读取行集合
                            String[] myLine = new String[iCount];   //单行对象
                            RangeSelector[] rsContentA = new RangeSelector[iCount];
                            for (j = 0; j < iCount; j++)
                            {
                                rsContentA[j] = new RangeSelector(lbContent.Items[j].ToString());
                            }
                            j = 0;
                            foreach (RangeSelector cont in rsContentA)
                            {
                                cont.acc();
                                point = cont.getCurPos();
                                myArray[m, j] = Convert.ToString(mySheet.Cells[point.Y, point.X].Value);    //不管什么类型都转为字符串
                                j++;
                                if (j > iCount) break;//xxxxxxx
                            }
                            myArray[m, j - 1] = DataTag;
                            m++;
                        }
                    }
                }
                //关闭当前活页簿
                myBook.Close();
                System.Windows.Forms.Application.DoEvents();
            }
            myExcel.Quit();
        }
        //字符串转坐标
        private Point pointPos(string strPos)
        {
            Point r = new Point(0, 0);
            char[] pa = strPos.ToUpper().ToCharArray();
            int i;
            for (i = 0; i < pa.Length; i++)
            {
                if (pa[i] >= 'A' && pa[i] <= 'Z')
                {
                    r.X = r.X * 26 + pa[i] - 'A' + 1;
                }
                else if (pa[i] >= '0' && pa[i] <= '9')
                {
                    r.Y = r.Y * 10 + pa[i] - '0';
                }
            }
            if (r.X == 0) r.X = 1;
            if (r.Y == 0) r.Y = 1;
            return r;
        }

        private void lbContent_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btClear_Click(object sender, EventArgs e)
        {
            lvFile.Clear();
            //myArray.Initialize();
        }

        private void tbAdd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnAdd_Click(sender, e);
            }
        }
        //将所有内容读取保存到myarray
        private void btReadWord_Click(object sender, EventArgs e)
        {

            int fileCount = lvFile.Items.Count;
            WordReader reader = new WordReader();
            reader.ResetData();
            for (int i = 0; i < fileCount; i++)
            {
                reader.ReadOne(lvFile.Items[i].SubItems[2].Text);
                reader.dataArray[i, reader.colCount] = lvFile.Items[i].SubItems[0].Text;
            }
            myArray = reader.dataArray;
            MessageBox.Show("读取完毕");
        }

        private void btReplace_Click(object sender, EventArgs e)
        {
            int fileCount = lvFile.Items.Count;
            WordReader reader = new WordReader();
            for (int i = 0; i < fileCount; i++)
            {
                reader.ReplaceOne(lvFile.Items[i].SubItems[2].Text, "代表推荐人选登记表", "代表候选人初步人选登记表");
            }
        }
        //批量修改excel
        private void btAutoEdit_Click(object sender, EventArgs e)
        {
            int fileCount = lvFile.Items.Count;
            ExcelReader reader = new ExcelReader();
            for (int i = 0; i < fileCount; i++)
            {
                reader.ReplaceOne(lvFile.Items[i].SubItems[2].Text);
            }
        }
    }

    class RangeSelector
    {
        public int pos;
        Point p1, p2;
        string rangestr;
        int width;
        int height;
        int type = 0; //0:未定义；1：单一单元格；2：同列单元格；3：区域单元格
        public RangeSelector()
        {
        }
        public RangeSelector(string s)
        {
            SetVal(s);
        }
        //重设起始位置
        public void SetStartVal(Point sp1)
        {
            p1 = sp1;
            type = 3;
            width = p2.X - p1.X + 1;
            height = p2.Y - p1.Y + 1;
            pos = 0;
        }
        public int getWidth()
        {
            return width;
        }
        public void SetVal(string s)
        {
            string s1, s2;
            rangestr = s;
            if (s == "")
            {
                s1 = s;
                type = 1;
                p1 = pointPos(s1);
                width = 0;
                height = 0;
            }
            if (s.Contains(":"))
            {
                int cp = s.IndexOf(":");
                s1 = s.Substring(0, cp);
                s2 = s.Substring(cp + 1);
                p1 = pointPos(s1);
                p2 = pointPos(s2);
                if (p1.X == p2.X) type = 2;
                else type = 3;
                width = p2.X - p1.X + 1;
                height = p2.Y - p1.Y + 1;
            }
            else
            {
                s1 = s;
                type = 1;
                p1 = pointPos(s1);
                width = 1;
                height = 1;
            }
            pos = 0;
        }
        //总大小
        public int Count()
        {
            return width * height;
        }
        //当前位置
        public Point getCurPos()
        {
            Point np = new Point();
            if (type == 1) return p1;
            if (type == 2)
            {
                np.X = p1.X;
                np.Y = p1.Y + pos % height;
                return np;
            }
            if (type == 3)
            {
                np.Y = p1.Y + (pos % (height * width)) / width;
                np.X = p1.X + pos % width;
                return np;
            }
            return np;
        }
        //移到下一格
        public bool acc()
        {
            pos++;
            if (pos >= Count()) return false;
            return true;
        }
        //移到下一行
        public bool lineacc()
        {
            pos += width;
            if (pos >= Count()) return false;
            return true;
        }
        //字符串转坐标
        private Point pointPos(string strPos)
        {
            Point r = new Point(0, 0);
            if (strPos == "") return r;
            char[] pa = strPos.ToUpper().ToCharArray();
            int i;
            for (i = 0; i < pa.Length; i++)
            {
                if (pa[i] >= 'A' && pa[i] <= 'Z')
                {
                    r.X = r.X * 26 + pa[i] - 'A' + 1;
                }
                else if (pa[i] >= '0' && pa[i] <= '9')
                {
                    r.Y = r.Y * 10 + pa[i] - '0';
                }
            }
            if (r.X == 0) r.X = 1;
            if (r.Y == 0) r.Y = 1;
            return r;
        }
    }
}
