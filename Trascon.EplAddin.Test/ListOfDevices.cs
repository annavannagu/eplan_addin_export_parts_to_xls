using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace Trascon.EplAddin.ExportToXLS
{
    class ListOfDevices
    {

        #region Var
        private string exportXMLTemp;
        private List<Part> devices;
        private string[] vs;
        private string temp;
        private XmlNodeList nodes;
        #endregion

        public ListOfDevices() { }

        public List<Part> GetAllDevices(string filepath)
        {
            exportXMLTemp = filepath;
            devices = new List<Part>();
            ImportXML();
            FillListfromXML();
            return devices;
        }


        private void ImportXML()
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(exportXMLTemp);
            
            XmlElement root = xmlDoc.DocumentElement;
            nodes = root.SelectNodes("device/part");
        }

        private void FillListfromXML()
        {
            foreach (XmlNode node in nodes)
            {
                bool isPartNo = node.Attributes["P_ARTICLE_DESCR2"] != null;
                bool isManufacturer = node.Attributes["P_ARTICLE_MANUFACTURER"] != null;
                bool isPlace1 = node.Attributes["P_DESIGNATION_FULLPLANT_WITHPREFIX"] != null;
                bool isPlace2 = node.Attributes["P_DESIGNATION_FULLLOCATION_WITHPREFIX"] != null;
                bool isDescription = node.Attributes["P_ARTICLE_DESCR1"] != null;
                bool isQuantity = node.Attributes["P_ARTICLE_QUANTITY_IN_PROJECT_UNIT"] != null;
                bool isHeader = node.Attributes["P_ARTICLE_PRODUCTGROUP"] != null;

                


                Part part = new Part();
                if (isPartNo)
                {
                    temp = node.Attributes["P_ARTICLE_DESCR2"].Value;
                    vs = temp.Split('@', ';');
                    int pos = -1;
                    pos = Array.IndexOf(vs, "??_??");
                    if (pos != -1) { vs[pos] = "ru_RU"; }                    
                    pos = Array.IndexOf(vs, "ru_RU");                                     
                    part.PartNo = vs[pos+1];
                }
                else { part.PartNo = "None"; };

                if (isManufacturer) { part.Manufacturer = node.Attributes["P_ARTICLE_MANUFACTURER"].Value; }
                else { part.Manufacturer = "None"; };

                if (isDescription)
                {
                    temp = node.Attributes["P_ARTICLE_DESCR1"].Value;
                    vs = temp.Split('@', ';');
                    int pos = -1;
                    pos = Array.IndexOf(vs, "??_??");
                    if (pos != -1) { vs[pos] = "ru_RU"; }
                    pos = Array.IndexOf(vs, "ru_RU");
                    part.Description = vs[pos + 1];
                }
                else { part.Description = "None"; };

                if (isQuantity)
                {
                    temp = node.Attributes["P_ARTICLE_QUANTITY_IN_PROJECT_UNIT"].Value;
                    double qnt = Convert.ToDouble(temp);
                    part.Quantity = qnt;
                }
                else { part.Quantity = 0; };

                if (isPlace1)
                {
                    if (isPlace2) { part.Place = node.Attributes["P_DESIGNATION_FULLPLANT_WITHPREFIX"].Value + node.Attributes["P_DESIGNATION_FULLLOCATION_WITHPREFIX"].Value; }
                    else { part.Place = node.Attributes["P_DESIGNATION_FULLPLANT_WITHPREFIX"].Value; }
                }
                else
                {
                    if (isPlace2) { part.Place = node.Attributes["P_DESIGNATION_FULLLOCATION_WITHPREFIX"].Value; }
                    else { part.Place = "None"; }
                }

                if (isHeader) { part.Header = node.Attributes["P_ARTICLE_PRODUCTGROUP"].Value; }

                if (isHeader && node.Attributes["P_ARTICLE_PRODUCTGROUP"].Value != "29")
                {
                    devices.Add(part);
                }
            }
        }        
    }
}
