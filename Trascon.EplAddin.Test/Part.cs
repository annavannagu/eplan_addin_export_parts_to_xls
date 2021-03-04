using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Trascon.EplAddin.ExportToXLS
{
    public class Part
    {
        public string PartNo { get; set; }

        public string Place { get; set; }

        public string Description { get; set; }

        public double Quantity { get; set; }

        public string Header { get; set; }

        public string Manufacturer { get; set; }
    }
}
