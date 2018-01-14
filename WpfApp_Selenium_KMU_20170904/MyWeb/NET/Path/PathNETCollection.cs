using System.Collections.Generic;
using WpfApp_Selenium_KMU_20170904.MyWeb.NET.Enums;


namespace WpfApp_Selenium_KMU_20170904.MyWeb.NET.Path
{
    internal class PathNETCollection : SortedDictionary<EnumPathNET, string>
    {
        internal PathNETCollection()
        {
            this.Add(EnumPathNET.EBT, "//div[@class='log']");
            this.Add(EnumPathNET.PITT, "//input[@type='pass' and @class='n']");
            this.Add(EnumPathNET.PHH, "//label[@for='0']");
            this.Add(EnumPathNET.LSAA, "//table[@id='Table']//child::tr[contains(@id,'Row')]");
            this.Add(EnumPathNET.TopCheckBox_Anketa, "//table[@id='Table']//child::tr[2]/td[1]");
            this.Add(EnumPathNET.EditSt_Anketa, "//td[@id='m']");
            this.Add(EnumPathNET.AOLA, "//td[@id='2']");
            this.Add(EnumPathNET.TabAdrAOLA, "//td[contains(@id,'C')]/div[3] //child::input[contains(@id,'w') and @type='text']");
            this.Add(EnumPathNET.TabAOLReggBT, "//*[contains(@id,'1')]");
            this.Add(EnumPathNET.TabAOLReggsList, "//*[contains(@id,'c')]/tbody/tr");
            this.Add(EnumPathNET.NextStt_Anketa, "//td[contains(@id,'T')]");
            this.Add(EnumPathNET.SaveDataStt_Anketa, "//td[@id='B']/div");
            this.Add(EnumPathNET.ChangePass_PersonData, "//*[contains(@id, '1')]");
            this.Add(EnumPathNET.PersonData, "//*[contains(@id, '0')]");
            this.Add(EnumPathNET.CollOfDoc_PersonData, "//table[contains(@id, 'T')]/tbody/tr");
            this.Add(EnumPathNET.SaveDataStt_PersonData, "//*[@id='D']");
        }
    }
}
