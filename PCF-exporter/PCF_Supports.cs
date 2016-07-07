using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using PCF_Functions;
using BuildingCoder;
using iv = PCF_Functions.InputVars;

namespace PCF_Functions
{

    public class SetSupportPipingSystem
    {
        public Result Execute(ExternalCommandData commandData, ref string msg, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;
            Application app = doc.Application;
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Transaction trans = new Transaction(doc, "Set System");
            ConnectorSet connectorSet = null;
            FamilyInstance familyInstance = null;
            MEPModel mepModel = null;
            Connector connectorToAdd = null;
            Connector systemDonorConnector = null;
            MEPSystem mepSystem = null;

            trans.Start();
            try
            {
                //Select Element to provide a system
                Element systemDonor = Util.SelectSingleElement(uidoc, "Select element in desired system.");

                if (systemDonor == null) throw new Exception("System assignment cancelled!");

                //Select Element to add to system
                Element elementToAdd = Util.SelectSingleElement(uidoc, "Select support to add to system.");

                if (elementToAdd == null) throw new Exception("System assignment cancelled!");

                switch (systemDonor.Category.Id.IntegerValue)
                {
                    case (int) BuiltInCategory.OST_PipeCurves:
                        Pipe pipe = (Pipe) systemDonor;
                        //Get connector set for the pipes
                        connectorSet = pipe.ConnectorManager.Connectors;
                        break;

                    case (int) BuiltInCategory.OST_PipeFitting:
                    case (int) BuiltInCategory.OST_PipeAccessory:
                        //Cast the element passed to method to FamilyInstance
                        familyInstance = (FamilyInstance) systemDonor;
                        //MEPModel of the elements is accessed
                        mepModel = familyInstance.MEPModel;
                        //Get connector set for the element
                        connectorSet = mepModel.ConnectorManager.Connectors;
                        break;
                        
                }

                //Get the connector
                if (connectorSet.IsEmpty) throw new Exception("No connectors in selected element. Select an element with connectors. Operation cancelled.");
                systemDonorConnector = (from Connector c in connectorSet where true select c).FirstOrDefault();
                mepSystem = systemDonorConnector.MEPSystem;

                Util.ErrorMsg("All tapping slots are taken. Manually delete unwanted values og increase number of tapping slots.");
                

                trans.Commit();

                //Debug
                //sbTaps.Append(tappedElement.LookupParameter(InputVars.PCF_ELEM_TAP1).AsString() == "");
                //sbTaps.AppendLine();
                //sbTaps.Append(tappedElement.LookupParameter(InputVars.PCF_ELEM_TAP2).AsString() + " " + tappedElement.LookupParameter(InputVars.PCF_ELEM_TAP2).AsString() == null);
                //sbTaps.AppendLine();
                //sbTaps.Append(tappedElement.LookupParameter(InputVars.PCF_ELEM_TAP3).AsString() + " " + tappedElement.LookupParameter(InputVars.PCF_ELEM_TAP3).AsString() == null);
                //sbTaps.AppendLine();



                //// Debugging
                //// Clear the output file
                //File.WriteAllBytes(InputVars.OutputDirectoryFilePath + "Taps.pcf", new byte[0]);

                //// Write to output file
                //using (StreamWriter w = File.AppendText(InputVars.OutputDirectoryFilePath + "Taps.pcf"))
                //{
                //    w.Write(sbTaps);
                //    w.Close();
                //}
            }

            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                trans.RollBack();
                return Result.Cancelled;
            }

            catch (Exception ex)
            {
                trans.RollBack();
                msg = ex.Message;
                return Result.Failed;
            }

            return Result.Succeeded;
        }
    }