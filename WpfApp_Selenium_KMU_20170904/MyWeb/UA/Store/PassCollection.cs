using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfApp_Selenium_KMU_20170904.MyWeb.Enums;

namespace WpfApp_Selenium_KMU_20170904.MyWeb.Path
{
    internal class PassCollection : SortedDictionary<Class_sss, string>
    {
        internal PassCollection()
        {
            this.Add(Class_sss.EAdm, "K");
            this.Add(Class_sss.ARr, "R");
            this.Add(Class_sss.Sss, "S");
        }
    }
}
