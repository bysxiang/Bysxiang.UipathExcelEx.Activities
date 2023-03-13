using System;
using System.Collections.Generic;
using System.Linq;
using System.Resources;
using System.Text;
using System.Threading.Tasks;

namespace Bysxiang.UipathExcelEx.Properties
{
    public static class Excel_Activities
    {
        private static ResourceManager rm = Resources.ResourceManager;

        public static string ExcelUsedRange_Name => rm.GetString(nameof(ExcelUsedRange_Name));

        public static string SizeInfo_Name => rm.GetString(nameof(SizeInfo_Name));

        public static string ExcelUsedRangeException => rm.GetString(nameof(ExcelUsedRangeException));
    }
}
