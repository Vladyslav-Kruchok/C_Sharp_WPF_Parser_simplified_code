using System.Collections.Generic;
using WpfApp_Selenium_KMU_20170904.CommonEnums;

namespace WpfApp_Selenium_KMU_20170904.CommonStore
{
    internal class LoginsCollection : SortedDictionary<WebSite, string>
    {
        internal LoginsCollection()
        {
            this.Add(WebSite.Xxx_ua, "xxx@xxx.ua");
        }
    }
}
