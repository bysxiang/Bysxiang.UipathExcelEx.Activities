using Bysxiang.UipathExcelEx.helpers;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UiPath.Excel;
using UiPath.Excel.Activities;

namespace Bysxiang.UipathExcelEx.Activities
{
    public abstract class ExcelExInteropActivity<T> : AsyncCodeActivity
    {
        protected bool CreateNew;

        [LocalizedCategory("Input")]
        [LocalizedDisplayName("SheetNameDisplayName")]
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; } = "Sheet1";

        protected ExcelExInteropActivity()
        {
            this.Constraints.Add(CheckParentConstraint.GetCheckParentConstraint<ExcelExInteropActivity<T>>(typeof(ExcelApplicationScope).Name));
        }

        // 以下代码从Uipath.Excel.Activities中复制，用以兼容不通的版本
        protected override IAsyncResult BeginExecute(AsyncCodeActivityContext context, AsyncCallback callback, 
            object state)
        {
            string sheetName = this.SheetName.Get((ActivityContext)context);
            WorkbookApplication workbook = context.DataContext.GetProperties()[UipathExcelHelper.GetWorkbookScopePropertyTag()].GetValue(context.DataContext) as WorkbookApplication;
            workbook.SetSheet(sheetName, this.CreateNew);
            Task<T> task = this.ExecuteAsync(context, workbook);
            TaskCompletionSource<T> tacs = new TaskCompletionSource<T>(state);
            Action<Task<T>> continuationAction = (Action<Task<T>>)(t =>
            {
                workbook.CloseSheet();
                if (t.IsFaulted)
                {
                    tacs.TrySetException(t.Exception.InnerExceptions);
                }    
                else if (t.IsCanceled)
                {
                    tacs.TrySetCanceled();
                }
                else
                {
                    tacs.TrySetResult(t.Result);
                }
                if (callback != null)
                {
                    callback(tacs.Task);
                }
            });
            task.ContinueWith(continuationAction);
            return tacs.Task;
        }

        protected override void EndExecute(AsyncCodeActivityContext context, IAsyncResult result)
        {
            Task<T> task = result as Task<T>;
            if (task.IsFaulted)
                throw task.Exception.InnerException;
            if (!task.IsCanceled)
            {
                if (!context.IsCancellationRequested)
                {
                    try
                    {
                        this.SetResult(context, task.Result);
                        return;
                    }
                    catch
                    {
                        context.MarkCanceled();
                        return;
                    }
                }
            }
            context.MarkCanceled();
        }

        protected abstract Task<T> ExecuteAsync(AsyncCodeActivityContext context, WorkbookApplication workbook);

        protected abstract void SetResult(AsyncCodeActivityContext context, T result);
    }
}
