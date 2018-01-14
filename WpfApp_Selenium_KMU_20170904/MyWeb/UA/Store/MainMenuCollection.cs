using System.Collections.Generic;
using WpfApp_Selenium_KMU_20170904.MyWeb.Enums;

namespace WpfApp_Selenium_KMU_20170904.MyWeb.Store
{
    internal class MainMenuCollection : SortedDictionary<MainMenu, string>
    {
        internal MainMenuCollection()
        {
            this.Add(MainMenu.Pro, "Н");
            this.Add(MainMenu.Lic, "Е");
            this.Add(MainMenu.Ent, "В");
            this.Add(MainMenu.OODoc, "З");
            this.Add(MainMenu.StEd, "З");
            this.Add(MainMenu.PI, "Ф");
            this.Add(MainMenu.Adm, "А");
        }
    }
}
