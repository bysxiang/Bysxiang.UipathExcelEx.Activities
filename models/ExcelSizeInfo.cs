using Bysxiang.UipathExcelEx.utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bysxiang.UipathExcelEx.models
{
    public class ExcelSizeInfo
    {
        public int Row { get; }

        public int Column { get; }

        public int RowCount { get; }

        public int ColumnCount { get; }

        public string ColumnName
        {
            get
            {
                return ExcelUtils.ToColumnName(Column);
            }
        }

        public string EndColumnName
        {
            get
            {
                return ExcelUtils.ToColumnName(Row + RowCount - 1);
            }
        }

        public string RangeStr
        {
            get
            {
                return string.Format("{0}{1}:{2}{3}", ColumnName, Row, EndColumnName, Row + RowCount - 1);
            }
        }
    }
}
