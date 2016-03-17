#region Header
//
// Copyright 2003-2015 by Autodesk, Inc. 
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted, 
// provided that the above copyright notice appears in all copies and 
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting 
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS. 
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC. 
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to 
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//
#endregion // Header

using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Linq;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Structure;

namespace PCF_Exporter
{

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    //[Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class App : IExternalApplication
    {
        //static Autodesk.Revit.DB.AddInId m_appId = new Autodesk.Revit.DB.AddInId(new Guid("709a7080-e6f5-49b4-810a-edd5bf5cb88d"));
        
        // get the absolute path of this assembly
        static string ExecutingAssemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        public Autodesk.Revit.UI.Result OnStartup(UIControlledApplication application)
        {
            AddMenu(application);
            return Autodesk.Revit.UI.Result.Succeeded;
        }

        public Autodesk.Revit.UI.Result OnShutdown(UIControlledApplication application)
        {
            return Autodesk.Revit.UI.Result.Succeeded;
        }
        
        private void AddMenu(UIControlledApplication application)
        {
            Autodesk.Revit.UI.RibbonPanel rvtRibbonPanel = application.CreateRibbonPanel("PCF Tools");
            PushButtonData data = new PushButtonData("PCFExporter","PCF Exporter",ExecutingAssemblyPath,"PCF_Exporter.FormCaller");
            PushButton pushButton = rvtRibbonPanel.AddItem(data) as PushButton;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class FormCaller : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                PCF_Exporter_form fm = new PCF_Exporter_form(commandData, message);
                fm.ShowDialog();
                
                fm.Close();
                return Result.Succeeded;
            }

            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }

            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
