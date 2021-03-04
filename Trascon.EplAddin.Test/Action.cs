using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using Eplan.EplApi.Base;

namespace Trascon.EplAddin.ExportToXLS
{
    public class Action : IEplAction
    {
        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            Name = "ActionTest";
            Ordinal = 20;
            return true;
        }                
        

        public bool Execute(ActionCallingContext oActionCallingContext)
        {
            Progress oProgress = new Progress("Состояние экспорта");
            oProgress.SetAllowCancel(true);

            #region Прогресс бар 1
            oProgress.BeginPart(60.0, "");
            oProgress.SetActionText("Получение данных из проекта..");
            oProgress.SetNeededSteps(1);
            oProgress.Step(1);
            #endregion

            //выбрать текущий проект
            SelectionSet Set = new SelectionSet();
            Project CurrentProject = Set.GetCurrentProject(true);

            //пути файлов
            string exportXMLPath = CurrentProject.ProjectDirectoryPath+@"\Temp.xml";
            string exportFilePath = CurrentProject.ProjectDirectoryPath + @"\DOC\Спецификация по шкафам.xlsx";

            //выгрузка в XML
            PartsService partsService = new PartsService();
            partsService.ExportPartsList(CurrentProject, exportXMLPath, 0);
            oProgress.EndPart(false);
            

            #region Прогресс бар 2
            oProgress.BeginPart(25.0, "");
            oProgress.SetActionText("Сортировка..");
            oProgress.SetNeededSteps(1);
            oProgress.Step(1);
            #endregion

            //загрузка в лист
            ListOfDevices listOfDevices = new ListOfDevices();            
            List<Part> devices = listOfDevices.GetAllDevices(exportXMLPath);

            var partsFiltered = from p in devices where ((p.PartNo != "None") && (p.Quantity != 0) && (p.PartNo != "Шильд")) orderby p.PartNo select p;
            oProgress.EndPart(false);

            #region Прогресс бар 3
            oProgress.BeginPart(15.0, "");
            oProgress.SetActionText("Экспорт в Excel..");
            oProgress.SetNeededSteps(1);
            oProgress.Step(1);
            #endregion

            XLSSerializer serializer = new XLSSerializer();
            serializer.Serialize(exportFilePath, partsFiltered.ToList());
            oProgress.EndPart(true); 
            
            return true;
        }

        public void GetActionProperties(ref ActionProperties actionProperties)
        {
        }
    }
}
