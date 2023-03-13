using Bysxiang.UipathExcelEx.utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UiPath.Excel.Helpers;

namespace Bysxiang.UipathExcelEx.models
{
    public class CellPosition
    {
        public int Row { get; set; }

        public int Column { get; set; }

        public CellPosition()
        {
            Row = 0;
            Column = 0;
        }

        public CellPosition(int row, int column)
        {
            Row = row;
            Column = column;
        }

        public bool IsValid => Row != 0 && Column != 0;

        public string ExcelRangeName()
        {
            return string.Format("{0}{1}", ExcelUtils.ToColumnName(Column), Row);
        }

        public override bool Equals(object obj)
        {
            return obj is CellPosition position &&
                   Row == position.Row &&
                   Column == position.Column;
        }

        public override int GetHashCode()
        {
            int hashCode = 240067226;
            hashCode = hashCode * -1521134295 + Row.GetHashCode();
            hashCode = hashCode * -1521134295 + Column.GetHashCode();
            return hashCode;
        }

    }
}
