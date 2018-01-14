using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfApp_Selenium_KMU_20170904.MyWeb.Enums;

namespace WpfApp_Selenium_KMU_20170904.MyWeb.Path
{
    internal class Class_sssCollection : SortedDictionary<Class_sss, string>
    {
        internal Class_sssCollection()
        {
            this.Add(Class_sss.EAdm, @"\");
            this.Add(Class_sss.ARr, @"\");
            this.Add(Class_sss.Sss, @"\");
        }
    }
}
