using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfApp_Selenium_KMU_20170904.MyWeb.Enums;

namespace WpfApp_Selenium_KMU_20170904.MyWeb.Path
{
    internal class PathContaintsTextCollection : SortedDictionary<ChoiceField, string>
    {
        internal PathContaintsTextCollection()
        {
            this.Add(ChoiceField.DataE, "Д");
            this.Add(ChoiceField.DataI, "Д");
            this.Add(ChoiceField.DiRrsP, "П");
            this.Add(ChoiceField.DiRrsN, "Пк");
        }
    }
}
