using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace snakeSkinV1
{
    public partial class Ribbon1
    {
        private Dictionary<Tuple<Excel.Range, Excel.Range>, Excel.Range> mainData = new Dictionary<Tuple<Excel.Range, Excel.Range>, Excel.Range>();

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

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
                    else { 
                    
                    }
                }
            }
            else
            {
                MessageBox.Show("擷取失敗");
            }
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
            //!important! TODO 使用了強制轉型
           Tuple<Excel.Range, Excel.Range > tmp = new Tuple<Excel.Range, Excel.Range>((Excel.Range)removeSelection.Items[0].Tag, (Excel.Range)removeSelection.Items[1].Tag);
           mainData[tmp] = (Excel.Range)removeSelection.Items[2].Tag;
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
            // string list a
            // int list b, c, d 
            List<String> a = new List<String>();
            List<String> tmp = new List<String>();
            List<int> b = new List<int>();
            List<int> c = new List<int>();
            List<double> d = new List<double>();

            foreach (var kvp in mainData)
            {
                var key = kvp.Key;
                var value = kvp.Value;
                //var keyItem1 = key.Item1.Address;
                //var keyItem2 = key.Item2.Address;
                //var worksheetof1 = key.Item1.Worksheet.Name;
                //var worksheetof2 = key.Item2.Worksheet.Name;
                //var worksheetof3 = value.Worksheet.Name;
                //sb.AppendLine($"Key: ([{worksheetof1}] {keyItem1}, [{worksheetof2}] {keyItem2}) => Value: [{worksheetof3}] {value.Address}");

                tmp.Add(key.Item1.Value2);
                a = a.Union(tmp).ToList();
                tmp.Add(key.Item2.Value2);
                a = a.Union(tmp).ToList();

                b.Add(a.FindIndex(var_important_coding_knowhow=> var_important_coding_knowhow==key.Item1.Value2));
                c.Add(a.FindIndex(var_important_coding_knowhow => var_important_coding_knowhow == key.Item2.Value2));
                d.Add(value.Value2);
            }

            string sa = string.Join(",", a.Select(x => x.ToString()));
            string sb = string.Join(",", b.Select(x => x.ToString()));
            string sc = string.Join(",", c.Select(x => x.ToString()));
            string sd = string.Join(",", d.Select(x => x.ToString()));

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "Rscript",
                Arguments = $"generate_sankey.R {sa} {sb} {sc} {sd}",
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

        private void listTest_Click(object sender, RibbonControlEventArgs e)
        {

            List<int> a = new List<int>();
            a.Add(1);
            a.Add(2);
            a.Add(3);
            string s = string.Join(",", a.Select(x => x.ToString()));
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
            MessageBox.Show(s);
            /**
             * todo:
             * 完成demo掛真實數據
             * 上b64
             */
        }

        private void todolist_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(
                "dc區增加'清除'按鈕\n" +
                "修改程式非阻擋式'圖表呈現'"
                );
        }
    }
}
