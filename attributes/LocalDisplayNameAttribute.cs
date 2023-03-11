using Bysxiang.UipathExcelEx.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bysxiang.UipathExcelEx.attributes
{
    internal class LocalDisplayNameAttribute : DisplayNameAttribute
    {
        public LocalDisplayNameAttribute(string displayName) : base(displayName)
        {
        }

        public override string DisplayName => Resources.ResourceManager.GetString(this.DisplayNameValue) ?? base.DisplayName;
    }
}
