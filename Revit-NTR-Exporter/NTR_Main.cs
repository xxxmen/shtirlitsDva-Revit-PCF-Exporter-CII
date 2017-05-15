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
        StringBuilder outputBuilder = new StringBuilder();

        public NTR_Exporter(ExternalCommandData cData)
        {
            ConfigurationData conf = new ConfigurationData(cData);

            outputBuilder.Append(conf._01_GEN);
            outputBuilder.Append(conf._02_AUFT);
            outputBuilder.Append(conf._03_TEXT);
            outputBuilder.Append(conf._04_LAST);
            outputBuilder.Append(conf._05_DN);
            outputBuilder.Append(conf._06_ISO);
        }
    }
}
