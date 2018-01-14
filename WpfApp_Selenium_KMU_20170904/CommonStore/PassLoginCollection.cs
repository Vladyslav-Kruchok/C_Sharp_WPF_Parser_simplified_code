using System.Collections.Generic;
using WpfApp_Selenium_KMU_20170904.CommonEnums;

namespace WpfApp_Selenium_KMU_20170904.CommonStore
{
    internal class PassLoginCollection : SortedDictionary<WebSite, string>
    {
        internal PassLoginCollection()
        {
            this.Add(WebSite.Xxx_ua, "Xxx");
            this.Add(WebSite.Xxx_net, "Xxx");
        }
    }
}
