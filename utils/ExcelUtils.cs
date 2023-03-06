using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bysxiang.UipathExcelEx.utils
{
    public static class ExcelUtils
    {
        /// <summary>
        /// 列序号转换为列名
        /// </summary>
        /// <param name="colNum"></param>
        /// <returns></returns>
        public static string ToColumnName(long colNum)
        {
            StringBuilder retVal = new StringBuilder();
            int x = 0;

            for (int n = (int)(Math.Log(25 * (colNum + 1)) / Math.Log(26)) - 1; n >= 0; n--)
            {
                x = (int)((Math.Pow(26, (n + 1)) - 1) / 25 - 1);
                if (colNum > x)
                {
                    retVal.Append(Convert.ToChar((int)(((colNum - x - 1) / Math.Pow(26, n)) % 26 + 65)));
                }
            }

            return retVal.ToString();
        }

        public static long ToColumnNum(string colName)
        {
            char[] chars = colName.ToUpper().ToCharArray();

            return (long)(Math.Pow(26, chars.Count() - 1)) *
                (Convert.ToInt32(chars[0]) - 64) +
                ((chars.Count() > 2) ? ToColumnNum(colName.Substring(1, colName.Length - 1)) :
                ((chars.Count() == 2) ? (Convert.ToInt32(chars[chars.Count() - 1]) - 64) : 0));
        }

    }
}
