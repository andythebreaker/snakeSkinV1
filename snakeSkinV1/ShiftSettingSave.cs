using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace snakeSkinV1
{
    internal class ShiftSettingSave
    {
        public ShiftSettingSave(string a, int b) {
            this.workSheetShiftNumber = b;
            this.workSheetName = a; 
        }
        public readonly string workSheetName;
        public readonly int workSheetShiftNumber;
        public string getTitle() { 
        return $"workingSheet:{this.workSheetName};值:{this.workSheetShiftNumber};";
        }
    }
}
