using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using NTR_Functions;

namespace NTR_Exporter
{
    class NTR_Exporter
    {
        public NTR_Exporter(ExternalCommandData cData)
        {
            ConfigurationData conf = new ConfigurationData(cData);
            
        }
    }
}
