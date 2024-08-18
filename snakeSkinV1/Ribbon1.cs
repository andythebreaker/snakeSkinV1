using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System.Reflection;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Microsoft.Office.Tools;

namespace snakeSkinV1
{
    public partial class Ribbon1
    {
        private Dictionary<Tuple<Excel.Range, Excel.Range>, Excel.Range> mainData = new Dictionary<Tuple<Excel.Range, Excel.Range>, Excel.Range>();
        private BrowserFormT1 maskMain = new BrowserFormT1();
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //https://akb48teamtp.fandom.com/zh-tw/wiki/AKB48_Team_TP%E6%88%90%E5%93%A1%E6%87%89%E6%8F%B4%E8%89%B2%E5%8F%8ACALL%E4%B8%80%E8%A6%BD%E8%A1%A8
            //a:source;b:target;c:value;1:background;2:text color
            c1.Color = System.Drawing.ColorTranslator.FromHtml("#ddb98b");
            c2.Color = System.Drawing.ColorTranslator.FromHtml("#000000");
            b1.Color = System.Drawing.ColorTranslator.FromHtml("#ffc0cb");
            b2.Color = System.Drawing.ColorTranslator.FromHtml("#008000");
            a1.Color = System.Drawing.ColorTranslator.FromHtml("#c4e1ff");
            a2.Color = System.Drawing.ColorTranslator.FromHtml("#bf4147");
            arrayColorSetSource1.Color= System.Drawing.ColorTranslator.FromHtml("#33ffff");
            arrayColorSetSource2.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
            arrayColorSetTarget1.Color = System.Drawing.ColorTranslator.FromHtml("#0000ff");
            arrayColorSetTarget2.Color = System.Drawing.ColorTranslator.FromHtml("#f1dd95");
            arrayColorSetData1.Color = System.Drawing.ColorTranslator.FromHtml("#feeeed");
            arrayColorSetData1.Color = System.Drawing.ColorTranslator.FromHtml("#7fc3ff");
            saveMirrorText.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            loadMirrorText.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        }

        private void displayData_Click(object sender, RibbonControlEventArgs e)
        {
            StringBuilder sb = new StringBuilder();

            foreach (var kvp in mainData)
            {
                var key = kvp.Key;
                var value = kvp.Value;
                var keyItem1 = key.Item1.Address;
                var keyItem2 = key.Item2.Address;
                var worksheetof1 = key.Item1.Worksheet.Name;
                var worksheetof2 = key.Item2.Worksheet.Name;
                var worksheetof3 = value.Worksheet.Name;

                sb.AppendLine($"Key: ([{worksheetof1}] {keyItem1}, [{worksheetof2}] {keyItem2}) => Value: [{worksheetof3}] {value.Address}");
            }

            MessageBox.Show(sb.ToString(), "Dictionary Data");
        }

        private void addMainData_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // Get the active Excel application
                Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");

                // Get the selected range of cells
                Excel.Range selectedRange = excelApp.Selection as Excel.Range;

                if (selectedRange != null)
                {
                    Excel.Range a = null;
                    Excel.Range b = null;
                    Excel.Range c = null;
                    foreach (Excel.Range cell in selectedRange.Cells)
                    {
                        if (a == null)
                        {
                            a = cell;
                        }
                        else if (b == null)
                        {
                            b = cell;
                        }
                        else if (c == null)
                        {
                            c = cell;
                        }
                        else
                        {
                            break;
                        }
                    }
                    var tmp = Tuple.Create(a, b);
                    mainData[tmp] = c;
                }
                else
                {
                    MessageBox.Show("No cells are selected.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private void sourceSelectMode_Click(object sender, RibbonControlEventArgs e)
        {
            targetSelectMode.Checked = false;
            valueSelectMode.Checked = false;
        }

        private void targetSelectMode_Click(object sender, RibbonControlEventArgs e)
        {
            sourceSelectMode.Checked = false;
            valueSelectMode.Checked = false;
        }

        private void valueSelectMode_Click(object sender, RibbonControlEventArgs e)
        {
            targetSelectMode.Checked = false;
            sourceSelectMode.Checked = false;
        }

        private void removeSelection_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            //NOTHING!
        }

        private void removeDC(int idx)
        {
            removeSelection.Items[idx].Tag = null;
            string tmp = idx == 0 ? "來源" : idx == 1 ? "目標" : "值";
            removeSelection.Items[idx].Label = $"「{tmp}」(尚未選取)";
        }

        private void doRemoveSelection_Click(object sender, RibbonControlEventArgs e)
        {
            removeDC(removeSelection.SelectedItemIndex);
        }

        private void capture_Click(object sender, RibbonControlEventArgs e)
        {
            var tmp = readUserSelectOne();
            if (isOne(tmp))
            {
                removeSelection.Items[sourceSelectMode.Checked ? 0 : targetSelectMode.Checked ? 1 : 2].Tag = tmp;
                removeSelection.Items[valueSelectMode.Checked ? 2 : targetSelectMode.Checked ? 1 : 0].Label = $"[{tmp.Worksheet.Name}]{tmp.Address}";
                removeSelection.SelectedItemIndex =
                    autoPreView.Checked ?
                    sourceSelectMode.Checked ? 0 : valueSelectMode.Checked ? 2 : 1
                    :
                    autoNextPT.Checked ? removeSelection.Items[0].Tag == null ? 0 : removeSelection.Items[1].Tag == null ? 1 : 2
                    :
                    removeSelection.SelectedItemIndex;
                if (autoNextPT.Checked)
                {
                    if (removeSelection.Items[0].Tag == null)
                    {
                        sourceSelectMode.Checked = true;
                        targetSelectMode.Checked = false;
                        valueSelectMode.Checked = false;
                        removeSelection.SelectedItemIndex =
                    autoPreView.Checked ? removeSelection.SelectedItemIndex :
                    0
                    ;
                    }
                    else if (removeSelection.Items[1].Tag == null)
                    {
                        sourceSelectMode.Checked = false;
                        targetSelectMode.Checked = true;
                        valueSelectMode.Checked = false;
                        removeSelection.SelectedItemIndex =
                    autoPreView.Checked ? removeSelection.SelectedItemIndex :
                    1
                    ;
                    }
                    else if (removeSelection.Items[2].Tag == null)
                    {
                        sourceSelectMode.Checked = false;
                        targetSelectMode.Checked = false;
                        valueSelectMode.Checked = true;
                        removeSelection.SelectedItemIndex =
                    autoPreView.Checked ? removeSelection.SelectedItemIndex :
                    2
                    ;
                    }
                    else
                    {

                    }
                }
            }
            else
            {
                MessageBox.Show("擷取失敗");
            }
        }

        public bool isOneColOrRow(Excel.Range r)
        {
            return isOneRow(r) ? true : isOneColumn(r) ? true : false;
        }

        public bool isOne(Excel.Range r)
        {
            //foreach (Excel.Range cell in r.Cells)
            //{
            if (r.Rows.Count == 1 && r.Columns.Count == 1)
            {
                // This cell represents a single cell
                return true;
            }
            else
            {
                // This is not a single cell (though it should always be in a Cells enumeration)
                return false;
            }
            //}
        }
        public bool isOneColumn(Excel.Range r)
        {
            if (r.Columns.Count == 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public bool isOneRow(Excel.Range r)
        {
            if (r.Rows.Count == 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public Excel.Range readUserSelectColOrRow()
        {
            try
            {
                // Get the active Excel application
                Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");

                // Get the selected range of cells
                Excel.Range selectedRange = excelApp.Selection as Excel.Range;

                if (selectedRange != null)
                {
                    if (isOneColOrRow(selectedRange))
                    {
                        return selectedRange;
                    }
                    else
                    {
                        MessageBox.Show("Too much, select only one row or column, thank you!");
                        return null;
                    }
                }
                else
                {
                    MessageBox.Show("No cells are selected.");
                    return null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
                return null;
            }
        }
        public Excel.Range readUserSelectOne()
        {
            try
            {
                // Get the active Excel application
                Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");

                // Get the selected range of cells
                Excel.Range selectedRange = excelApp.Selection as Excel.Range;

                if (selectedRange != null)
                {
                    if (isOne(selectedRange))
                    {
                        return selectedRange;
                    }
                    else
                    {
                        MessageBox.Show("Too much, select only one, thank you!");
                        return null;
                    }
                }
                else
                {
                    MessageBox.Show("No cells are selected.");
                    return null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
                return null;
            }
        }

        private void readUserSelectB_Click(object sender, RibbonControlEventArgs e)
        {
            var tmp = readUserSelectOne();
            MessageBox.Show(tmp == null ? "return nul" : tmp.Address);
        }

        private void addOne_Click(object sender, RibbonControlEventArgs e)
        {
            if (safe3.Checked) {
                if (removeSelection.Items[0].Tag == null || removeSelection.Items[1].Tag == null || removeSelection.Items[2].Tag == null) {
                    MessageBox.Show("未選擇滿3個格子!");
                }
            }
            //!important! TODO 使用了強制轉型
            Tuple<Excel.Range, Excel.Range> tmp = new Tuple<Excel.Range, Excel.Range>((Excel.Range)removeSelection.Items[0].Tag, (Excel.Range)removeSelection.Items[1].Tag);
            mainData[tmp] = (Excel.Range)removeSelection.Items[2].Tag;
            if (hotfixAutoReset41.Checked) {
                removeDC(0); removeDC(1); removeDC(2);
                sourceSelectMode.Checked = true;
                targetSelectMode.Checked = false;
                valueSelectMode.Checked = false;
            }
        }

        private void addRibbonDropdownItemB_Click(object sender, RibbonControlEventArgs e)
        {
            //Process.Start("C:\\Users\\ai\\Documents\\andy\\code\\tmp\\p\\y\\bin\\Debug\\y.exe");
            // Create a new process start info
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "Rscript",
                Arguments = "generate_sankey.R a,b,c 1,2 2,3 3,4",
                WorkingDirectory = @"C:\Users\ai\Documents\andy\code\snakeskin\masterR",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            // Start the process
            using (Process process = Process.Start(startInfo))
            {
                // Capture and display the output
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                // Display the output
                if (!string.IsNullOrEmpty(output))
                {
                    Console.WriteLine("Output: " + output);
                }

                // Display the error
                if (!string.IsNullOrEmpty(error))
                {
                    Console.WriteLine("Error: " + error);
                }
            }
        }

        private void processData_Click(object sender, RibbonControlEventArgs e)
        {
            List<String> a = new List<String>();
            List<String> tmp = new List<String>();
            List<int> b = new List<int>();
            List<int> c = new List<int>();
            List<double> d = new List<double>();

            foreach (var kvp in mainData)
            {
                var key = kvp.Key;
                var value = kvp.Value;

                tmp.Add(key.Item1.Value2);
                a = a.Union(tmp).ToList();
                tmp.Add(key.Item2.Value2);
                a = a.Union(tmp).ToList();

                b.Add(a.FindIndex(var_important_coding_knowhow => var_important_coding_knowhow == key.Item1.Value2));
                c.Add(a.FindIndex(var_important_coding_knowhow => var_important_coding_knowhow == key.Item2.Value2));
                try
                {
                    d.Add(Convert.ToDouble(value.Value2));
                }
                catch (InvalidCastException var_error)
                {
                    MessageBox.Show("[錯誤!] 這是一個錯誤，旨在表明「儲存格(" + value.Worksheet.Name + ")" + value.Address +
                        "」並不是實數。 \n提醒:這個儲存格必須要是實數(整數或小數)!\n相關資訊:這個出錯的儲存格表述了「"
                        + key.Item1.Value2 +
                         "到" +
                        key.Item2.Value2
                         + "」的轉換關係；並且他的值是"
                        + "「" + value.Value2 +
                        "」。\n狀態:「出圖」動作並未完成請修改excel工作表中的值後再重新「出圖」。\n其他錯誤資訊:" +
                        var_error.ToString());
                    return;
                }
                catch (FormatException var_error)
                {
                    MessageBox.Show("[錯誤!] 這是一個錯誤，旨在表明「儲存格(" + value.Worksheet.Name + ")" + value.Address +
                        "」並不是實數。 \n提醒:這個儲存格必須要是實數(整數或小數)!\n相關資訊:這個出錯的儲存格表述了「"
                        + key.Item1.Value2 +
                         "到" +
                        key.Item2.Value2
                         + "」的轉換關係；並且他的值是"
                        + "「" + value.Value2 +
                        "」。\n狀態:「出圖」動作並未完成請修改excel工作表中的值後再重新「出圖」。\n其他錯誤資訊:" +
                        var_error.ToString());
                    return;
                }
                catch (OverflowException var_error)
                {
                    MessageBox.Show("[錯誤!] 這是一個錯誤，旨在表明「儲存格(" + value.Worksheet.Name + ")" + value.Address +
                        "」並不是實數。 \n提醒:這個儲存格必須要是實數(整數或小數)!\n相關資訊:這個出錯的儲存格表述了「"
                        + key.Item1.Value2 +
                         "到" +
                        key.Item2.Value2
                         + "」的轉換關係；並且他的值是"
                        + "「" + value.Value2 +
                        "」。\n狀態:「出圖」動作並未完成請修改excel工作表中的值後再重新「出圖」。\n其他錯誤資訊:" +
                        var_error.ToString());
                    return;
                }
            }

            string sa = string.Join(",", a.Select(x => Convert.ToBase64String(Encoding.UTF8.GetBytes(x))));
            string sb = string.Join(",", b.Select(x => x.ToString()));
            string sc = string.Join(",", c.Select(x => x.ToString()));
            string sd = string.Join(",", d.Select(x => x.ToString()));

            string tempPath = Path.GetTempPath();
            string fileName = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            string file_extension = ".txt";
            string file_name_with_extention = fileName + file_extension;
            string filePath = Path.Combine(tempPath, file_name_with_extention);
            string content = $"{sa}\n{sb}\n{sc}\n{sd}\n";
            File.WriteAllText(filePath, content);

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "Rscript",
                Arguments = $"generate_sankey_via_file.R {filePath}",
                WorkingDirectory = @"C:\Users\ai\Documents\andy\code\snakeskin\masterR",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            // Start the process
            using (Process process = Process.Start(startInfo))
            {
                // Capture and display the output
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                // Display the output
                if (!string.IsNullOrEmpty(output))
                {
                    File.WriteAllText(Path.Combine(tempPath, fileName + "output" + file_extension), output);
                }

                // Display the error
                if (!string.IsNullOrEmpty(error))
                {
                    File.WriteAllText(Path.Combine(tempPath, fileName + "error" + file_extension), error);
                }
            }

        }

        private void listTest_Click(object sender, RibbonControlEventArgs e)
        {
            // Get the path of the temporary directory
            string tempPath = Path.GetTempPath();

            // Generate the file name based on the current date and time
            string fileName = DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".txt";

            // Combine the path and the file name
            string filePath = Path.Combine(tempPath, fileName);

            // Define the content to write to the file
            string content = "Hello, this is a test file.";

            // Write the content to the file
            File.WriteAllText(filePath, content);
            //List<int> a = new List<int>();
            //a.Add(1);
            //a.Add(2);
            //a.Add(3);
            //string s = string.Join(",", a.Select(x => x.ToString()));
            //List<int> b = new List<int>();
            //List<int> c = new List<int>();
            //a.Add(1);
            //b.Add(2);
            //c.Add(2);
            //c = c.Union(a).ToList();
            //c = c.Union(b).ToList();
            //StringBuilder sb = new StringBuilder();
            //foreach (var item in c)
            //{
            //    sb.Append(item.ToString());
            //    sb.Append("; ");
            //}
            //MessageBox.Show(s);
        }

        private void todolist_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(
                "dc區增加'清除'按鈕\n" +
                "修改程式非阻擋式'圖表呈現'的說明(只要先開瀏覽器)\n"
                );
        }

        private void addCell(string locationStr, string dataStr, double dataInt)
        {
            if (dataStr == null)
            {
                Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                Excel.Range firstRow = excelApp.get_Range(locationStr);
                //firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                //Excel.Range newFirstRow = excelApp.get_Range("A1");
                firstRow.Value2 = dataInt;
            }
            else if (dataInt.Equals(double.NaN))
            {
                Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                Excel.Range firstRow = excelApp.get_Range(locationStr);
                //firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                //Excel.Range newFirstRow = excelApp.get_Range("A1");
                firstRow.Value2 = dataStr;
            }
            else
            {
                MessageBox.Show("內部程式錯誤(addCell)這個函式僅適用於開發環境");
            }
        }

        private void writeMainDataDumb_Click(object sender, RibbonControlEventArgs e)
        {
            addCell("B1", "person1", double.NaN); addCell("C1", "person2", double.NaN); addCell("A2", "person3", double.NaN); addCell("A3", "person4", double.NaN);
            addCell("B2", null, 5); addCell("C2", null, 6); addCell("B3", null, 2); addCell("C3", null, 3);
            //         person1 person2 
            //person3    5        6
            //person4    2        3
        }

        private void galleryNumTest_Click(object sender, RibbonControlEventArgs e)
        {
            //Microsoft.Office.Interop.Excel.Range x = readUserSelectOne();
            //MessageBox.Show($"{x.Interior.Color}\n{x.Font.Color}");
            MessageBox.Show(clearVisual.Tag == null ? "t" : "f");
        }

        private void editData_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonGallery gallery = (RibbonGallery)sender;
            RibbonDropDownItem selectedItem = gallery.SelectedItem;

            if (modeEdit.SelectedItem.Tag.ToString() == "view")
            {
                clearVisual_do();
                System.Tuple<Microsoft.Office.Interop.Excel.Range, Microsoft.Office.Interop.Excel.Range> k = (System.Tuple<Microsoft.Office.Interop.Excel.Range, Microsoft.Office.Interop.Excel.Range>)selectedItem.Tag;
                Microsoft.Office.Interop.Excel.Range stuffToChangeColor = mainData[k];//!important! 強轉型
                a1.Tag = k.Item1.Interior.Color;
                a2.Tag = k.Item1.Font.Color;
                b1.Tag = k.Item2.Interior.Color;
                b2.Tag = k.Item2.Font.Color;
                c1.Tag = stuffToChangeColor.Interior.Color;
                c2.Tag = stuffToChangeColor.Font.Color;
                //stuffToChangeColor Style background color = #ddb98b text color = #ffc0cb
                // Set the background color to #ddb98b
                stuffToChangeColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(c1.Color);
                // Set the text color to #ffc0cb
                stuffToChangeColor.Font.Color = System.Drawing.ColorTranslator.ToOle(c2.Color);
                k.Item1.Interior.Color = System.Drawing.ColorTranslator.ToOle(a1.Color);
                k.Item1.Font.Color = System.Drawing.ColorTranslator.ToOle(a2.Color);
                k.Item2.Interior.Color = System.Drawing.ColorTranslator.ToOle(b1.Color);
                k.Item2.Font.Color = System.Drawing.ColorTranslator.ToOle(b2.Color);
                clearVisual.Tag = k;
            }
            else if (modeEdit.SelectedItem.Tag.ToString() == "del")
            {
                mainData.Remove((System.Tuple<Microsoft.Office.Interop.Excel.Range, Microsoft.Office.Interop.Excel.Range>)selectedItem.Tag);//!important! 強轉型
            }
            else
            {
                MessageBox.Show("[Error] Warning: The program has reached an unexpected logic section. This action was not properly executed. Please contact the developer for further assistance.\r\n ");
            }
        }

        private void loadShiftSetting() {
            Dictionary<string, int> old = new Dictionary<string, int>();//!important!不能有兩個名稱一樣的活頁簿
            foreach (RibbonDropDownItem i in shiftSetting.Items)
            {
                old[((ShiftSettingSave)i.Tag).workSheetName] = ((ShiftSettingSave)i.Tag).workSheetShiftNumber;
            }

            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Sheets sh = excelApp.Sheets;
            List<RibbonDropDownItem> ui2append = new List<RibbonDropDownItem>();
            foreach (Microsoft.Office.Interop.Excel.Worksheet i in sh)
            {
                //MessageBox.Show(i.Name);
                RibbonDropDownItem editData_tmp = this.Factory.CreateRibbonDropDownItem();
                try
                {
                    ShiftSettingSave sht = new ShiftSettingSave(i.Name, old[i.Name]);
                    editData_tmp.Label = sht.getTitle();
                    editData_tmp.Tag = sht;
                    ui2append.Add(editData_tmp);
                }
                catch (Exception e)
                {
                    ShiftSettingSave sht = new ShiftSettingSave(i.Name, 0);
                    editData_tmp.Label = sht.getTitle();
                    editData_tmp.Tag = sht;
                    ui2append.Add(editData_tmp);
                }
            }
            shiftSetting.Items.Clear();
            foreach (var i in ui2append)
            {
                shiftSetting.Items.Add(i);
            }
        }

        private void updateWorkSheetShiftSetting(object sender, RibbonControlEventArgs e)
        {
            loadShiftSetting();
        }

        private void editDataLoad(object sender, RibbonControlEventArgs e)
        {
            editData.Items.Clear();
            foreach (var d in mainData)
            {
                RibbonDropDownItem editData_tmp = this.Factory.CreateRibbonDropDownItem();
                editData_tmp.Label = $"來源:{d.Key.Item1.Value2};目標:{d.Key.Item2.Value2};值:{d.Value.Value2};";
                editData_tmp.Tag = d.Key;
                editData_tmp.ScreenTip = $"來源:[{d.Key.Item1.Worksheet.Name}]{d.Key.Item1.Address};目標:[{d.Key.Item2.Worksheet.Name}]{d.Key.Item2.Address};值:[{d.Value.Worksheet.Name}]{d.Value.Address};";
                editData.Items.Add(editData_tmp);
            }
        }

        private void modeEdit_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void a1show_Click(object sender, RibbonControlEventArgs e)
        {
            a1.ShowDialog();
        }

        private void a2show_Click(object sender, RibbonControlEventArgs e)
        {
            a2.ShowDialog();
        }

        private void b1show_Click(object sender, RibbonControlEventArgs e)
        {
            b1.ShowDialog();
        }

        private void b2show_Click(object sender, RibbonControlEventArgs e)
        {
            b2.ShowDialog();
        }

        private void c1show_Click(object sender, RibbonControlEventArgs e)
        {
            c1.ShowDialog();
        }

        private void c2show_Click(object sender, RibbonControlEventArgs e)
        {
            c2.ShowDialog();
        }



        private void setCellsColor(
            System.Tuple<Microsoft.Office.Interop.Excel.Range, Microsoft.Office.Interop.Excel.Range> k,
            double color_a1,
            double color_a2,
            double color_b1,
            double color_b2,
            double color_c1,
            double color_c2
            )
        {
            k.Item1.Interior.Color = color_a1;
            k.Item1.Font.Color = color_a2;
            k.Item2.Interior.Color = color_b1;
            k.Item2.Font.Color = color_b2;
            mainData[k].Interior.Color = color_c1;
            mainData[k].Font.Color = color_c2;
        }

        private void clearVisual_do()
        {
            if (clearVisual.Tag != null)
            {
                setCellsColor((System.Tuple<Microsoft.Office.Interop.Excel.Range, Microsoft.Office.Interop.Excel.Range>)clearVisual.Tag,
                    (double)a1.Tag,
                    (double)a2.Tag,
                    (double)b1.Tag,
                    (double)b2.Tag,
                    (double)c1.Tag,
                    (double)c2.Tag);

            }
        }

        private void clearVisual_Click(object sender, RibbonControlEventArgs e)
        {
            clearVisual_do();
        }
        enum typeSourceTargetData
        {
            source,target,data
        }
        private Tuple<List<double>, List<double>> savePrvColor(Excel.Range r,typeSourceTargetData tSTD) {
            List<double> a = new List<double>();
            List<double> b = new List<double>();
            System.Drawing.Color c1,c2;
            switch (tSTD)
            {
                case typeSourceTargetData.source:
                    c1 = arrayColorSetSource1.Color;
                    c2 = arrayColorSetSource2.Color;
                    break;
                case typeSourceTargetData.target:
                    c1 = arrayColorSetTarget1.Color;
                    c2 = arrayColorSetTarget2.Color;
                    break;
                case typeSourceTargetData.data:
                    c1 = arrayColorSetData1.Color;
                    c2 = arrayColorSetData2.Color;
                    break;
                default:
                    c1= System.Drawing.ColorTranslator.FromHtml("#000000");
                    c2 = System.Drawing.ColorTranslator.FromHtml("#000000");
                    break;
            }
            foreach (Range c in r.Cells)
            {
                a.Add(c.Interior.Color);
                c.Interior.Color = System.Drawing.ColorTranslator.ToOle(c1);
                b.Add(c.Font.Color);
                c.Font.Color = System.Drawing.ColorTranslator.ToOle(c2);
            }
            return Tuple.Create(a, b);
        }
        private void arraySetSource_Click(object sender, RibbonControlEventArgs e)
        {
            arraySetSource.Tag = readUserSelectColOrRow();
            //if neq null change color
            if (arraySetSource.Tag != null&&displayColorAfterSelect.Checked) {
                Tuple<List<double>, List<double>> savePrvColor_obj = savePrvColor((Excel.Range)arraySetSource.Tag,typeSourceTargetData.source);
                arrayColorSetSource1.Tag= savePrvColor_obj.Item1;
                arrayColorSetSource2.Tag= savePrvColor_obj.Item2;
            }
        }

        private void arraySetTarget_Click(object sender, RibbonControlEventArgs e)
        {//if neq null change color
            arraySetTarget.Tag = readUserSelectColOrRow();
            if (arraySetTarget.Tag != null && displayColorAfterSelect.Checked)
            {
                Tuple<List<double>, List<double>> savePrvColor_obj = savePrvColor((Excel.Range)arraySetTarget.Tag, typeSourceTargetData.target);
                arrayColorSetTarget1.Tag = savePrvColor_obj.Item1;
                arrayColorSetTarget2.Tag = savePrvColor_obj.Item2;
            }
        }

        private void arraySetData_Click(object sender, RibbonControlEventArgs e)
        {
            if (arraySetData.Tag==null) {
                MessageBox.Show("error! you did not select your data! action not finish!");
                return;
            }
            Tuple<List<Excel.Range>, List<Excel.Range>, List<Excel.Range>> previewData = (Tuple<List<Excel.Range>, List<Excel.Range>, List<Excel.Range>>)arraySetData.Tag;
            for (int i = 0; i < previewData.Item3.Count; i++)
            {
                Range c = previewData.Item3[i];
                int a = i / previewData.Item2.Count;//source
                int b = i % previewData.Item2.Count;//target
                Excel.Range a_c = previewData.Item1[a];
                Excel.Range b_c = previewData.Item2[b];
                Tuple<Excel.Range, Excel.Range> tmp = new Tuple<Excel.Range, Excel.Range>(a_c, b_c);
                mainData[tmp] = c;
                if (arrayColorSetData1.Tag != null) {
                    c.Interior.Color = ((List<double>)arrayColorSetData1.Tag)[i];
                }
                if (arrayColorSetData2.Tag != null)
                {
                    c.Font.Color = ((List<double>)arrayColorSetData2.Tag)[i];
                }
            }
            if (arrayColorSetSource1.Tag != null) {
                for(int  i = 0; i < previewData.Item1.Count; i++)
                {
                    (previewData.Item1[i]).Interior.Color = ((List<double>)arrayColorSetSource1.Tag)[i];
                }
            }
            if (arrayColorSetSource2.Tag != null)
            {
                for (int i = 0; i < previewData.Item1.Count; i++)
                {
                    (previewData.Item1[i]).Font.Color = ((List<double>)arrayColorSetSource2.Tag)[i];
                }
            }
            if (arrayColorSetTarget1.Tag != null)
            {
                for (int i = 0; i < previewData.Item2.Count; i++)
                {
                    (previewData.Item2[i]).Interior.Color = ((List<double>)arrayColorSetTarget1.Tag)[i];
                }
            }
            if (arrayColorSetTarget2.Tag != null)
            {
                for (int i = 0; i < previewData.Item2.Count; i++)
                {
                    (previewData.Item2[i]).Font.Color = ((List<double>)arrayColorSetTarget2.Tag)[i];
                }
            }
            arraySetData.Tag = null;
        }

        private void previewArray_Click(object sender, RibbonControlEventArgs e)
        { 
            if (((Excel.Range)arraySetSource.Tag).Count==0|| ((Excel.Range)arraySetTarget.Tag).Count==0) {
                MessageBox.Show("error, you did not select array source or array tatget! action not finish!");
                return;
            }
            Excel.Range d = readUserSelectOne().Resize[((Excel.Range)arraySetSource.Tag).Count, ((Excel.Range)arraySetTarget.Tag).Count];
            Tuple<List<double>, List<double>> savePrvColor_obj = savePrvColor(d, typeSourceTargetData.data);
            arrayColorSetData1.Tag = savePrvColor_obj.Item1;
            arrayColorSetData2.Tag = savePrvColor_obj.Item2;
            //d.Interior.Color = System.Drawing.ColorTranslator.ToOle(a1.Color);
            List<Excel.Range> s = new List<Excel.Range>();
            List<Excel.Range> t = new List<Excel.Range>();
            List<Excel.Range> d_list = new List<Excel.Range>();
            foreach (Range c in ((Excel.Range)arraySetSource.Tag).Cells)
            {
                s.Add(c);
            }
            foreach (Range c in ((Excel.Range)arraySetTarget.Tag).Cells)
            {
                t.Add(c);
            }
            foreach (Range c in d.Cells)
            {
                d_list.Add(c);
            }
            Tuple<List<Excel.Range>, List<Excel.Range>, List<Excel.Range>> tmp = new Tuple<List<Excel.Range>, List<Excel.Range>, List<Excel.Range>>(
            s, t, d_list
                );
            arraySetData.Tag = tmp;
        }

        private void picColor1_Click(object sender, RibbonControlEventArgs e)
        {
            arrayColorSetSource1.ShowDialog();
        }

        private void picColor2_Click(object sender, RibbonControlEventArgs e)
        {
            arrayColorSetSource2.ShowDialog();
        }

        private void picColor3_Click(object sender, RibbonControlEventArgs e)
        {
            arrayColorSetTarget1.ShowDialog();
        }

        private void picColor4_Click(object sender, RibbonControlEventArgs e)
        {
            arrayColorSetTarget2.ShowDialog();
        }

        private void picColor5_Click(object sender, RibbonControlEventArgs e)
        {
            arrayColorSetData1.ShowDialog();
        }

        private void picColor6_Click(object sender, RibbonControlEventArgs e)
        {
            arrayColorSetData2.ShowDialog();
        }

        private void rainbowTest_Click(object sender, RibbonControlEventArgs e)
        {
            // Define colors
            double white = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("#ddb98b"));
            double green = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("#ffc0cb"));
            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            // Loop to change colors
            bool toggle = true;
            for (int i = 0; i < 10; i++) // Example loop, adjust as needed
            {
                excelApp.Range["A1"].Interior.Color = toggle ? green : white;
                excelApp.Range["A2"].Interior.Color = toggle ? white : green;
                toggle = !toggle;
                Thread.Sleep(500); // Wait for 0.5 seconds
            }

        }

        private async void rainbowMG_Click(object sender, RibbonControlEventArgs e)
        {
            foreach (var d in mainData)
            {
                double color_tmp_1 = d.Key.Item1.Interior.Color;
                double color_tmp_2 = d.Key.Item1.Font.Color;
                double color_tmp_3 = d.Key.Item2.Interior.Color;
                double color_tmp_4 = d.Key.Item2.Font.Color;
                double color_tmp_5 = d.Value.Interior.Color;
                double color_tmp_6 = d.Value.Font.Color;
                setCellsColor(d.Key, 
                    System.Drawing.ColorTranslator.ToOle(a1.Color)
                    , System.Drawing.ColorTranslator.ToOle(a2.Color)
                    , System.Drawing.ColorTranslator.ToOle(b1.Color)
                    , System.Drawing.ColorTranslator.ToOle(b2.Color)
                    , System.Drawing.ColorTranslator.ToOle(c1.Color)
                    , System.Drawing.ColorTranslator.ToOle(c2.Color));
                await Task.Delay(500); // Non-blocking delay
                setCellsColor(d.Key, color_tmp_1, color_tmp_2, color_tmp_3, color_tmp_4, color_tmp_5, color_tmp_6);
            }
        }

        private void testsave_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            excelApp.ActiveWorkbook.CustomDocumentProperties.Add("testP1",false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeNumber,48);
        }

        private void testloadsave_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            int mustbe48 = excelApp.ActiveWorkbook.CustomDocumentProperties["testP1"].Value;
            MessageBox.Show(mustbe48.ToString());
        }

        private void worksheetcodenametest_Click(object sender, RibbonControlEventArgs e)
        {
            Range x = readUserSelectOne();
            string y = x.Worksheet.CodeName;
            MessageBox.Show(x.Value2);

        }

       public List<string> SplitString(string str, int maxChunkSize)
        {
            List<string> result = new List<string>();
            if (str.Length < maxChunkSize) { 
            result.Add(str);
            }
            else
            {
                for (int i = 0; i < str.Length; i += maxChunkSize)
                {
                    // Ensure we do not exceed the string length
                    if (i + maxChunkSize > str.Length)
                    {
                        result.Add(str.Substring(i));
                    }
                    else
                    {
                        result.Add(str.Substring(i, maxChunkSize));
                    }
                }
            }
            return result;
        }

        private void saveMap_Click(object sender, RibbonControlEventArgs e)
        {
            editData.Items.Clear();//!important!從edit data load 複製過來的
            List<DicSave> mirror = new List<DicSave>();
            foreach (var d in mainData)
            {
                mirror.Add(new DicSave(
                d.Key.Item1.Worksheet.Name,d.Key.Item1.Address,d.Key.Item2.Worksheet.Name,d.Key.Item2.Address,d.Value.Worksheet.Name,d.Value.Address
                ));
            }
            string jsonStr = JsonConvert.SerializeObject(mirror);
            //MessageBox.Show(jsonStr);
            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            // excelApp.ActiveWorkbook.Variables{ "mainDataMirror" } = 
            List<string> jsonStr255 = SplitString(jsonStr, 255);
            foreach (var (js255, index) in jsonStr255.Select((value, i) => (value, i)))
            {
                // Use 'js255' and 'index' here
            excelApp.ActiveWorkbook.CustomDocumentProperties.Add($"mainDataMirror{index}", false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, js255);

            }
            excelApp.ActiveWorkbook.CustomDocumentProperties.Add("mainDataMirrorLength", false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeNumber, jsonStr255.Count);
        }

        private void loadMap_Click(object sender, RibbonControlEventArgs e)
        {
            if (emptyWhenLoad.Checked)
            {
                mainData.Clear();
            }
            /*
[{"source":{"worksheet":"工作表1","address":"$A$2"},"target":{"worksheet":"工作表1","address":"$B$1"},"value":{"worksheet":"工作表1","address":"$B$2"}},{"source":{"worksheet":"工作表1","address":"$A$2"},"target":{"worksheet":"工作表1","address":"$C$1"},"value":{"workshe
             */
            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            string jsonStr = "";//excelApp.ActiveWorkbook.CustomDocumentProperties["mainDataMirror"].Value;
            int mainDataMirrorLength = excelApp.ActiveWorkbook.CustomDocumentProperties["mainDataMirrorLength"].Value;
            for (int i = 0; i < mainDataMirrorLength; i++) {
                string js255 = excelApp.ActiveWorkbook.CustomDocumentProperties[$"mainDataMirror{i}"].Value;
                jsonStr += js255;
            }
            List<DicSave> mirror = JsonConvert.DeserializeObject<List<DicSave>>(jsonStr);//?? new List<DicSave>();
            foreach (DicSave d in mirror)
            {
                Excel.Workbook workbook = excelApp.ActiveWorkbook;
                Excel.Worksheet worksheet1 = workbook.Sheets[d.source.worksheet];
                Excel.Range range1 = worksheet1.get_Range(d.source.address);
                Excel.Worksheet worksheet2 = workbook.Sheets[d.target.worksheet];
                Excel.Range range2 = worksheet2.get_Range(d.target.address);
                Excel.Worksheet worksheet3 = workbook.Sheets[d.value.worksheet];
                Excel.Range range3 = worksheet3.get_Range(d.value.address);
                var tmp = Tuple.Create(range1,range2);
                mainData[tmp] = range3;
            }
        }

        private void exportMap_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            string jsonStr = "";
            int mainDataMirrorLength = excelApp.ActiveWorkbook.CustomDocumentProperties["mainDataMirrorLength"].Value;
            for (int i = 0; i < mainDataMirrorLength; i++)
            {
                string js255 = excelApp.ActiveWorkbook.CustomDocumentProperties[$"mainDataMirror{i}"].Value;
                jsonStr += js255;
            }
            if (saveMirrorText.ShowDialog() == DialogResult.OK)
            {
                // Get the selected file path
                string filePath = saveMirrorText.FileName;

                // Write the string to the file
                File.WriteAllText(filePath, jsonStr);

                // Optionally, you can display a message to the user
                MessageBox.Show($"File saved successfully at: {filePath}");
            }
        }

        private int shiftSettingQuery(IList<RibbonDropDownItem> listOfItems,string toFind) {
            foreach (var item in listOfItems)
            {
                if (((ShiftSettingSave)item.Tag).workSheetName==toFind)
                {
                    return ((ShiftSettingSave)item.Tag).workSheetShiftNumber;
                }
            }
            MessageBox.Show($"Error! Not found! Can't find worksheet : {toFind} in shift setting list. Using shift number = 0.");
            return 0;
        }

        private void importMap_Click(object sender, RibbonControlEventArgs e)
        {
            if (ableShift.Checked)
            { 
            loadShiftSetting();
            }
            if (emptyWhenLoad.Checked)
            {
                mainData.Clear();
            }
            if (loadMirrorText.ShowDialog() == DialogResult.OK)
            {
                // Get the selected file path
                string filePath = loadMirrorText.FileName;

                // Read the content of the file
                string jsonStr = File.ReadAllText(filePath);
                Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                List<DicSave> mirror = JsonConvert.DeserializeObject<List<DicSave>>(jsonStr);//這段是從loadMap_Click複製過來的
                foreach (DicSave d in mirror)
                {
                    Excel.Workbook workbook = excelApp.ActiveWorkbook;
                 Excel.Worksheet worksheet1 = workbook.Sheets[d.source.worksheet];
                    Excel.Worksheet worksheet2 = workbook.Sheets[d.target.worksheet];
                    Excel.Worksheet worksheet3 = workbook.Sheets[d.value.worksheet];
                    if (ableShift.Checked) {
                        Excel.Range range1 = worksheet1.get_Range(d.source.address);
                        Excel.Range range2 = worksheet2.get_Range(d.target.address);
                        Excel.Range range3 = worksheet3.get_Range(d.value.address);
                        var tmp = Tuple.Create(ShiftRange( range1, shiftSettingQuery(shiftSetting.Items,range1.Worksheet.Name)),
                           ShiftRange(range2, shiftSettingQuery(shiftSetting.Items, range2.Worksheet.Name)));
                        mainData[tmp] = ShiftRange(range3, shiftSettingQuery(shiftSetting.Items, range3.Worksheet.Name));
                    }
                    else { 
                    Excel.Range range1 = worksheet1.get_Range(d.source.address);
                    Excel.Range range2 = worksheet2.get_Range(d.target.address);
                    Excel.Range range3 = worksheet3.get_Range(d.value.address);
                    var tmp = Tuple.Create(range1, range2);
                    mainData[tmp] = range3;
                    
                    }
                    
                }

            }
        }

        private void shiftSetting_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonGallery gallery = (RibbonGallery)sender;
            RibbonDropDownItem selectedItem = gallery.SelectedItem;
            var tmp = readUserSelectColOrRow().Cells ;
            if (tmp == null)
            {
                //MessageBox.Show("action not finish!");
                return;
            }
            else
            {
                //MessageBox.Show(tmp.Count.ToString()+"~"+ selectedItem.Label);
                selectedItem.Tag = new ShiftSettingSave(((ShiftSettingSave)selectedItem.Tag).workSheetName,
                    ((tmp.Count-1)<0)?0: tmp.Count-1) ;
            }
        }

        public static Excel.Range ShiftRange(Excel.Range range, int shiftDown)
        {
            return range.Cells[1+shiftDown, 1];//~~植樹問題所以應該是1+shiftdown-1~~
        }

        private void testActivateWindows_Click(object sender, RibbonControlEventArgs e)
        {
            ShiftRange(readUserSelectOne(),3).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
        }

        private void addSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            Task.Run(() =>
            {
                maskMain.ShowDialog();
            });
        }

        private void testAddRow_Click(object sender, RibbonControlEventArgs e)
        {
            maskMain.AddRow(readUserSelectOne());
        }

        private void newWindowsTag_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
