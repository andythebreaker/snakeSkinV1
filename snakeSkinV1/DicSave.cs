using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
//using Newtonsoft.Json;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Web;

namespace snakeSkinV1
{/*public class CustomDictionary : IEnumerable<KeyValuePair<string, int>>
{
    private string jsonData = "{}";

    public int this[string key]
    {
        get
        {
            var dict = JsonConvert.DeserializeObject<Dictionary<string, int>>(jsonData);
            if (dict != null && dict.TryGetValue(key, out int value))
            {
                return value;
            }
            throw new KeyNotFoundException($"Key '{key}' not found.");
        }
        set
        {
            var dict = JsonConvert.DeserializeObject<Dictionary<string, int>>(jsonData) ?? new Dictionary<string, int>();
            dict[key] = value;
            jsonData = JsonConvert.SerializeObject(dict);
        }
    }

    public override string ToString()
    {
        return jsonData;
    }

    public IEnumerator<KeyValuePair<string, int>> GetEnumerator()
    {
        var dict = JsonConvert.DeserializeObject<Dictionary<string, int>>(jsonData);
        return dict?.GetEnumerator() ?? new Dictionary<string, int>().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }*/
    internal class DicSave //: IEnumerable<KeyValuePair<Tuple<Excel.Range, Excel.Range>, Excel.Range>>
    {
        /*public Dictionary<Tuple<Excel.Range, Excel.Range>, Excel.Range> thisDictionary()
        {
            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            int SSkeysCount = excelApp.ActiveWorkbook.CustomDocumentProperties["SSkeysCount"].Value;
            //SSskey...num = string ...source
            //SStkey...num = string ...target
            //SSvkey...num = string ...value
            Dictionary<Tuple<Excel.Range, Excel.Range>, Excel.Range> stuff2return = new Dictionary<Tuple<Excel.Range, Excel.Range>, Excel.Range>();
            for (int i = 0; i < SSkeysCount; i++) {
                string sk = "SSskey" + i.ToString();
                string tk = "SStkey" + i.ToString();
                string vk = "SSvkey" + i.ToString();
                string s = excelApp.ActiveWorkbook.CustomDocumentProperties[sk].Value;
                string t = excelApp.ActiveWorkbook.CustomDocumentProperties[tk].Value;
                string v = excelApp.ActiveWorkbook.CustomDocumentProperties[vk].Value;
                //TODO!notyet excelApp.Worksheets
            }
                return new Dictionary<Tuple<Excel.Range, Excel.Range>, Excel.Range>();
        }

        private IEnumerator<KeyValuePair<Tuple<Excel.Range, Excel.Range>, Excel.Range>> GetEnumerator()
        {
            try
            {
                return thisDictionary().GetEnumerator();
            }
            catch (Exception e)
            {
                MessageBox.Show("DicSave.cs:IEnumerator" + e.Message);
                throw e;
            }
        }

        IEnumerator<KeyValuePair<Tuple<Excel.Range, Excel.Range>, Excel.Range>> IEnumerable<KeyValuePair<Tuple<Excel.Range, Excel.Range>, Excel.Range>>.GetEnumerator()
        {
            return GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
            //throw new NotImplementedException();
        }*/
        public struct worksheet_and_address
        {
            public string worksheet;
            public string address;
        }

        public worksheet_and_address source { get; set; }
        public worksheet_and_address target { get; set; }
        public worksheet_and_address value { get; set; }

        public DicSave(string source_worksheet,
            string source_address,
            string target_worksheet,
            string target_address,
            string value_worksheet,
            string value_address) {
            source = new worksheet_and_address
            {
                worksheet = source_worksheet,
                address = source_address
            };

            target = new worksheet_and_address
            {
                worksheet = target_worksheet,
                address = target_address
            };

            value = new worksheet_and_address
            {
                worksheet = value_worksheet,
                address = value_address
            };
        }
    }
}
