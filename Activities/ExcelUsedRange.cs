using Bysxiang.UipathExcelEx.attributes;
using Bysxiang.UipathExcelEx.helpers;
using Bysxiang.UipathExcelEx.models;
using Bysxiang.UipathExcelEx.Properties;
using Microsoft.Office.Interop.Excel;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using UiPath.Excel;
using UiPath.Excel.Activities;
using up = UiPath.Excel;

namespace Bysxiang.UipathExcelEx.Activities
{
    [LocalDisplayName("ExcelUsedRange_Name")]
    public sealed class ExcelUsedRange : ExcelExInteropActivity<ExcelSizeInfo>
    {
        [LocalizedCategory("Output")]
        public OutArgument<ExcelSizeInfo> SizeInfo { get; set; }

        protected override Task<ExcelSizeInfo> ExecuteAsync(AsyncCodeActivityContext context, up.WorkbookApplication workbook)
        {
            Range range = null;
            try
            {
                range = workbook.CurrentWorksheet.UsedRange;
                return Task.Run(() =>
                {
                    ExcelSizeInfo sizeInfo = new ExcelSizeInfo(range);
                    return sizeInfo;
                });
            }
            catch (COMException ex)
            {
                throw new up.ExcelException(string.Format(Excel_Activities.ExcelUsedRangeException));
            }
            finally
            {
                ComHelpers.ReleaseAndClearComObject(ref range);
            }
        }

        protected override void SetResult(AsyncCodeActivityContext context, ExcelSizeInfo result)
        {
            this.SetResult(context, result);
        }
    }
}
