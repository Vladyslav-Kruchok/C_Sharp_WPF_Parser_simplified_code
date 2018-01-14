using System.Collections.Generic;
using WpfApp_Selenium_KMU_20170904.MyWeb.Enums;

namespace WpfApp_Selenium_KMU_20170904.MyWeb.Path
{
    internal class SPSP : SortedDictionary<ChoiceField, string>
    {
        internal SPSP()
        {
            this.Add(ChoiceField.DataE, "following-sibling::div[1]/div/div[10]/div[3]/*//input");
            this.Add(ChoiceField.DataI, "following-sibling::div[1]/div/div[10]/div[1]/*//input");
            this.Add(ChoiceField.DiRrsP, "following-sibling::div[1]/div/div[26]/*//input");
            this.Add(ChoiceField.DiRrsN, "following-sibling::div[1]/div/div[26]/*//input");
        }
    }
}
