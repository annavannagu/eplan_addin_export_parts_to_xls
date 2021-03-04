using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Gui;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Trascon.EplAddin.ExportToXLS
{
    public class AddInModule : IEplAddIn
    {
        public bool OnRegister(ref bool bLoadOnStart)
        {
            bLoadOnStart = true;
            return true;
        }

        public bool OnUnregister()
        {
            return true;
        }

        public bool OnInit()
        {
            return true;
        }

        public bool OnInitGui() //добавление Item в меню Eplan
        {
            Menu OurMenu = new Menu();
            OurMenu.AddMainMenu("Экспорт в Excel", Menu.MainMenuName.eMainMenuUtilities, "Выгрузить все изделия", "ActionTest",
                "Экспорт всех изделий проекта по шкафам для расчетов", 1);            
            return true;
        }

        public bool OnExit()
        {
            return true;
        }
    }
}
