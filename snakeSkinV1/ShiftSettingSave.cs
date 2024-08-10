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
            this.workSheetShiftNumber = b;//~~ b-1<0?0:b-1;//植樹問題所以應該是1+shiftdown-1~~
            this.workSheetName = a; 
        }
        public readonly string workSheetName;
        public readonly int workSheetShiftNumber;
        public string getTitle() { 
        return $"workingSheet:{this.workSheetName};值:{this.workSheetShiftNumber};";
        }
    }
}
