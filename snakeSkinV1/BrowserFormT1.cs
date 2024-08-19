using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace snakeSkinV1
{
    public partial class BrowserFormT1 : Form
    {
        private ColorDialog cF;
        private ColorDialog cB;
        private StatusStrip bottom;
        private ToolStripDropDownButton debugButton;
        private ToolStripMenuItem testComWmainToolStripMenuItem;
        private ToolStripDropDownButton colorGroup;
        private ToolStripMenuItem frontC;
        private ToolStripMenuItem backC;
        private ToolStripMenuItem testTag;
        private DataGridView dataGridView;

        public struct imm
        {
            public bool isMaskedMain;
            public bool item1new;
            public bool item2new;
            public bool errorCode;
        }

        public struct maskDo {
            public imm inn;
            public int item1idx;
            public int item2idx;
            public double val;
            public KeyValuePair<Tuple<Excel.Range, Excel.Range>, Excel.Range> keyValuePair;
        }

        public struct abcd {
            public List<String> a;
            public List<int> b;
            public List<int> c;
            public List<double> d;
        }

        public List<string> maskDoA=new List<string>();
        public List<maskDo> md=new List<maskDo>();

        public BrowserFormT1()
        {
            InitializeComponent();

            maskDoA = new List<string>();
            md = new List<maskDo>();

            dataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = true
            };

            dataGridView.Columns.Add("TextColumn", "位置:活頁表");//row id 1
            dataGridView.Columns.Add("TextColumn", "位置:地址");//row id 2
            dataGridView.Columns.Add("TextColumn", "資料輔助紀錄:前景(請使用者忽略這個行的資料)");//row id 3
            dataGridView.Columns.Add("TextColumn", "資料輔助紀錄:背景(請使用者忽略這個行的資料)");//row id 4
            dataGridView.Columns.Add("TextColumn", "檢視紀錄");

            var buttonColumn = new DataGridViewButtonColumn
            {
                Name = "ButtonColumn",
                HeaderText = "閃爍檢視",
                Text = "點擊我在Excel主程式中檢視格子的位置",
                UseColumnTextForButtonValue = true
            };
            dataGridView.Columns.Add(buttonColumn);

            // Fill DataGridView with some data
            /*  for (int i = 0; i < 5; i++)
              {
                  dataGridView.Rows.Add($"Row {i + 1}", $"This is text for Row {i + 1}", "abc");
              }*/

            dataGridView.CellClick += DataGridView_CellClick;

            // Add DataGridView to the form
            this.Controls.Add(dataGridView);

            // Set form properties
            this.Width = 800;
            this.Height = 600;

            // Handle FormClosing event
            this.FormClosing += BrowserFormT1_FormClosing;
        }

        public void clearMem() { 
        maskDoA.Clear();
            md.Clear();
        }
        
        private void DataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dataGridView.Columns["ButtonColumn"].Index && e.RowIndex >= 0)
            {
                // MessageBox.Show($"Button clicked in row {dataGridView.Rows[ e.RowIndex].Tag}!");

                ((Excel.Range)(dataGridView.Rows[e.RowIndex].Tag)).Interior.Color = System.Drawing.ColorTranslator.ToOle(cB.Color);
                ((Excel.Range)(dataGridView.Rows[e.RowIndex].Tag)).Font.Color = System.Drawing.ColorTranslator.ToOle(cF.Color);
                (dataGridView.Rows[e.RowIndex]).Cells[5].Value = true;
            }
        }

        private void BrowserFormT1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Prompt the user to confirm closing
            // var result = MessageBox.Show("Are you sure you want to close?", "Confirm Close", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            // if (result == DialogResult.No)
            // {
            // Cancel the close operation
            //      e.Cancel = true;
            //  }

            foreach (DataGridViewRow r in dataGridView.Rows)
            {
                if (((bool)(r.Cells[5].Value)) == true)
                {
                    ((Excel.Range)r.Tag).Interior.Color = r.Cells[4].Value;
                    ((Excel.Range)r.Tag).Font.Color = r.Cells[3].Value;
                    r.Cells[5].Value = false;
                }
            }
        }

        // Method to add a new row
        public void AddRow(Excel.Range r)
        {
            string notInit = "notInit";
            dataGridView.Rows.Add(notInit, notInit, notInit, notInit, notInit, false);
            dataGridView.Rows[dataGridView.Rows.Count - 1].Tag = r;
        }

        public abcd backbrust(List<String> a,
        List<int> b,
        List<int> c,
        List<double> d){
            List<String> aNew = new List<String>(a);
            aNew.AddRange(maskDoA);
            foreach (var i in md)
            {
                if (!i.inn.item1new)
                {
                    b.Add(a.FindIndex(var_important_coding_knowhow => var_important_coding_knowhow == i.keyValuePair.Key.Item1.Value2));
                }
                else { 
                b.Add(a.Count+i.item1idx-1);
                }
                if (!i.inn.item2new)
                {
                    c.Add(a.FindIndex(var_important_coding_knowhow => var_important_coding_knowhow == i.keyValuePair.Key.Item2.Value2));
                }
                else
                {
                    c.Add(a.Count + i.item2idx - 1);
                }
                d.Add(i.val);
            }
            abcd abcdTmp= new abcd();   
            abcdTmp.a = aNew;
            abcdTmp.b = b;
            abcdTmp.c= c;
            abcdTmp.d = d;
            return abcdTmp; 
        }

        public imm isMasked(KeyValuePair< Tuple<Excel.Range, Excel.Range>,Excel.Range> tr3)
        {
            refreshTable();
            imm immTmp=new imm();
            immTmp.errorCode = false;
            maskDo maskDoTmp=new maskDo();
            foreach (DataGridViewRow r in dataGridView.Rows)
            {
                if (
                //(
                ((string)r.Cells[1].Value == tr3.Key.Item1.Worksheet.Name)
                &&
               ((string)r.Cells[2].Value == tr3.Key.Item1.Address)
                    ) { 
                immTmp.item1new = true;
                    maskDoA.Add(tr3.Key.Item1.Value2);
                    maskDoTmp.item1idx=maskDoA.Count;
                }//||
                       if  (
                  ((string)r.Cells[1].Value == tr3.Key.Item2.Worksheet.Name)
                &&
              ((string)r.Cells[2].Value == tr3.Key.Item2.Address)
                    )//)
                    { immTmp.item2new = true;
                    maskDoA.Add(tr3.Key.Item2.Value2);
                    maskDoTmp.item2idx = maskDoA.Count;
                }
                /*{
                    immTmp.isMaskedMain
                    return true;
                }*/
            }
            immTmp.isMaskedMain = (immTmp.item1new || immTmp.item2new) ? true : false;
            maskDoTmp.inn = immTmp;
            try
            {//這一段是複製過來的
                maskDoTmp.val= Convert.ToDouble(tr3.Value.Value2);
            }
            catch (InvalidCastException var_error)
            {
                MessageBox.Show("[錯誤!] 這是一個錯誤，旨在表明「儲存格(" + tr3.Value.Worksheet.Name + ")" + tr3.Value.Address +
                "」並不是實數。 \n提醒:這個儲存格必須要是實數(整數或小數)!\n相關資訊:這個出錯的儲存格表述了「"
                    + tr3.Key.Item1.Value2 +
                     "到" +
                    tr3.Key.Item2.Value2
                + "」的轉換關係；並且他的值是"
                    + "「" + tr3.Value.Value2 +
                    "」。\n狀態:「出圖」動作並未完成請修改excel工作表中的值後再重新「出圖」。\n其他錯誤資訊:" +
                    var_error.ToString());
                immTmp.errorCode = true;
            }
            catch (FormatException var_error)
            {
                MessageBox.Show("[錯誤!] 這是一個錯誤，旨在表明「儲存格(" + tr3.Value.Worksheet.Name + ")" + tr3.Value.Address +
                "」並不是實數。 \n提醒:這個儲存格必須要是實數(整數或小數)!\n相關資訊:這個出錯的儲存格表述了「"
                    + tr3.Key.Item1.Value2 +
                     "到" +
                    tr3.Key.Item2.Value2
                     + "」的轉換關係；並且他的值是"
                    + "「" + tr3.Value.Value2 +
                    "」。\n狀態:「出圖」動作並未完成請修改excel工作表中的值後再重新「出圖」。\n其他錯誤資訊:" +
                    var_error.ToString());
                immTmp.errorCode = true;
            }
            catch (OverflowException var_error)
            {
                MessageBox.Show("[錯誤!] 這是一個錯誤，旨在表明「儲存格(" + tr3.Value.Worksheet.Name + ")" + tr3.Value.Address +
                "」並不是實數。 \n提醒:這個儲存格必須要是實數(整數或小數)!\n相關資訊:這個出錯的儲存格表述了「"
                    + tr3.Key.Item1.Value2 +
                     "到" +
                    tr3.Key.Item2.Value2
                     + "」的轉換關係；並且他的值是"
                    + "「" + tr3.Value.Value2 +
                    "」。\n狀態:「出圖」動作並未完成請修改excel工作表中的值後再重新「出圖」。\n其他錯誤資訊:" +
                    var_error.ToString());
                immTmp.errorCode = true;
            }
            maskDoTmp.keyValuePair = tr3;
            if(immTmp.isMaskedMain) md.Add(maskDoTmp);

            return immTmp;
        }

        private void InitializeComponent()
        {
            //這裡的東西是IDE加的
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BrowserFormT1));
            this.cF = new System.Windows.Forms.ColorDialog();
            this.cB = new System.Windows.Forms.ColorDialog();
            this.bottom = new System.Windows.Forms.StatusStrip();
            this.colorGroup = new System.Windows.Forms.ToolStripDropDownButton();
            this.backC = new System.Windows.Forms.ToolStripMenuItem();
            this.frontC = new System.Windows.Forms.ToolStripMenuItem();
            this.debugButton = new System.Windows.Forms.ToolStripDropDownButton();
            this.testComWmainToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.testTag = new System.Windows.Forms.ToolStripMenuItem();
            this.bottom.SuspendLayout();
            this.SuspendLayout();
            // 
            // bottom
            // 
            this.bottom.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.colorGroup,
            this.debugButton});
            this.bottom.Location = new System.Drawing.Point(0, 303);
            this.bottom.Name = "bottom";
            this.bottom.Size = new System.Drawing.Size(362, 22);
            this.bottom.TabIndex = 0;
            this.bottom.Text = "statusStrip1";
            this.bottom.Click += new System.EventHandler(this.statusStrip1_Click);
            // 
            // colorGroup
            // 
            this.colorGroup.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.colorGroup.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.backC,
            this.frontC});
            this.colorGroup.Image = global::snakeSkinV1.Properties.Resources.colorpalette;
            this.colorGroup.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.colorGroup.Name = "colorGroup";
            this.colorGroup.Size = new System.Drawing.Size(29, 20);
            this.colorGroup.Text = "colorGroup";
            // 
            // backC
            // 
            this.backC.Name = "backC";
            this.backC.Size = new System.Drawing.Size(180, 22);
            this.backC.Text = "檢視背景色設定";
            this.backC.Click += new System.EventHandler(this.backC_Click);
            // 
            // frontC
            // 
            this.frontC.Name = "frontC";
            this.frontC.Size = new System.Drawing.Size(180, 22);
            this.frontC.Text = "檢視前景色設定";
            this.frontC.Click += new System.EventHandler(this.frontC_Click);
            // 
            // debugButton
            // 
            this.debugButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.debugButton.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.testComWmainToolStripMenuItem,
            this.testTag});
            this.debugButton.Image = global::snakeSkinV1.Properties.Resources.bugbeetle;
            this.debugButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.debugButton.Name = "debugButton";
            this.debugButton.Size = new System.Drawing.Size(29, 20);
            this.debugButton.Text = "除錯用按鈕";
            this.debugButton.Click += new System.EventHandler(this.debugButton_Click);
            // 
            // testComWmainToolStripMenuItem
            // 
            this.testComWmainToolStripMenuItem.Name = "testComWmainToolStripMenuItem";
            this.testComWmainToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.testComWmainToolStripMenuItem.Text = "testComWmain";
            this.testComWmainToolStripMenuItem.Click += new System.EventHandler(this.testComWmainToolStripMenuItem_Click);
            // 
            // testTag
            // 
            this.testTag.Name = "testTag";
            this.testTag.Size = new System.Drawing.Size(180, 22);
            this.testTag.Text = "testTag";
            // 
            // BrowserFormT1
            // 
            this.ClientSize = new System.Drawing.Size(362, 325);
            this.Controls.Add(this.bottom);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "BrowserFormT1";
            this.Text = "重複名稱遮罩控制面板";
            this.Load += new System.EventHandler(this.BrowserFormT1_Load);
            this.bottom.ResumeLayout(false);
            this.bottom.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

            //這裡的東西是我加的andythebreaker
            cF.Color = System.Drawing.ColorTranslator.FromHtml("#4C592E");
            cB.Color = System.Drawing.ColorTranslator.FromHtml("#7A9FBF");
        }

        private void testComWmainToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            double green = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("#ffc0cb"));
            excelApp.Range["A1"].Interior.Color = green;
        }

        private void refreshTable() { 
         foreach (DataGridViewRow r in dataGridView.Rows)
            {
                r.Cells[0].Value = ((Excel.Range)r.Tag).Value2;
                r.Cells[1].Value = ((Excel.Range)r.Tag).Worksheet.Name;
                r.Cells[2].Value = ((Excel.Range)r.Tag).Address;
                r.Cells[4].Value = ((Excel.Range)r.Tag).Interior.Color;
                r.Cells[3].Value = ((Excel.Range)r.Tag).Font.Color;
            }
        }

        private void BrowserFormT1_Load(object sender, EventArgs e)
        {
            refreshTable();
        }

        private void debugButton_Click(object sender, EventArgs e)
        {
            // MessageBox.Show("這個功能是給開發者除錯用的，除非你知道你在幹嘛，不然不要擅自進來這個區域，很危險的XD");
        }

        private void backC_Click(object sender, EventArgs e)
        {
            cB.ShowDialog();
        }

        private void frontC_Click(object sender, EventArgs e)
        {
            cF.ShowDialog();
        }

        private void statusStrip1_Click(object sender, EventArgs e)
        {
            dataGridView.Rows[0].Tag = "this is a test tag";
            dataGridView.Rows[0].Cells[2].Value = "bear";//0->1->2
        }
    }
}
