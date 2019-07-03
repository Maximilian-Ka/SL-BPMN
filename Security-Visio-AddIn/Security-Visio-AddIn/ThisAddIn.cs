using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;

namespace Security_Visio_AddIn
{
    public partial class ThisAddIn
    {
        // TODO: Validator-Methoden in Validator-Klasse auslagern
        // TODO: Weitere Ausnahmen behandeln.
        // TODO: Issue Handling implementieren
        // TODO: Von Validator zu Validator kann die übergebene Liste an Shapes gekürzt werden, damit Shapes nicht immer wieder überprüft werden.
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string docPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) + @"\test\myDrawing.vsdx";
            Visio.Document doc = this.Application.Documents.Open(docPath);
            Visio.Shapes vsoShapes = getShapesFromPage();
            Visio.ValidationRuleSet gatewayValidatorRuleSet = doc.Validation.RuleSets.Add("Gateway Validation");
            gatewayValidatorRuleSet.Description = "Verify that the gateways are correctly used in the document.";
            Application.RuleSetValidated += new Visio.EApplication_RuleSetValidatedEventHandler(HandleRuleSetValidatedEvent);          
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void HandleRuleSetValidatedEvent(Visio.ValidationRuleSet RuleSet)
        {
            gatewayValidator(getShapesFromPage(), getActiveDocument());
            //if (RuleSet.Name == "Gateway Validation")
            //{
            //    gatewayValidator(getShapesFromPage(), getActiveDocument());
            //}
            //if (RuleSet.Name == "Inspection Validation")
            //{
            //    inspectionValidator(getShapesFromPage(), getActiveDocument());
            //}
            //if (RuleSet.Name == "Violation Validation")
            //{
            //    violationValidator(getShapesFromPage(), getActiveDocument());
            //}
            //if (RuleSet.Name == "Surveillance Validation")
            //{
            //    surveillanceValidator(getShapesFromPage(), getActiveDocument());
            //}

        }


        public void insertRuleSet(Visio.Document document)
        {
            Visio.ValidationRuleSet gatewayValidatorRuleSet = document.Validation.RuleSets.Add("Gateway Validation");
            gatewayValidatorRuleSet.Description = "Verify that the gateways are correctly used in the document.";

            Visio.ValidationRule customRule1 = gatewayValidatorRuleSet.Rules.Add("distinctFlows2sequenceFlow");
            customRule1.Category = "Gateway";
            customRule1.Description = "Wenn ungleiche Sequenzflüsse zusammengeführt werden, muss der ausgehende Sequenzfluss DangerFlow sein";

            Visio.ValidationRule customRule2 = gatewayValidatorRuleSet.Rules.Add("equalFlows2distinctFlow");
            customRule2.Category = "Gateway";
            customRule2.Description = "Eingehende Sequenzflüsse sind ungleich dem ausgehenden Sequenzfluss. Muss gleich sein.";
            //
            //
            Visio.ValidationRuleSet inspectionValidatorRuleSet = document.Validation.RuleSets.Add("Inspection Validation");
            inspectionValidatorRuleSet.Description = "Verify that the Inspection-Shapes are correctly used in the document.";

            Visio.ValidationRule customRule3 = inspectionValidatorRuleSet.Rules.Add("glued2DshapesMissing");
            customRule3.Category = "inspection-shape";
            customRule3.Description = "Kein 2D-Shape an das Inspection-Shape geklebt.";

            Visio.ValidationRule customRule4 = inspectionValidatorRuleSet.Rules.Add("gluedViolationEvent");
            customRule4.Category = "inspection-shape";
            customRule4.Description = "Kein Violation-Event an das Inspections-Shape geklebt.";
            //
            //Template
            //Visio.ValidationRuleSet ****ValidatorRuleSet = document.Validation.RuleSets.Add("Name");
            //*****ValidatorRuleSet.Description = "";
            //Visio.ValidationRule customRule* = inspectionValidatorRuleSet.Rules.Add("");
            //customRule*.Category = "";
            //customRule*.Description = "";
        }



        public Visio.Document getActiveDocument()
        {
            Visio.Document doc = Application.ActiveDocument;
            return doc;
        }

        public Visio.Shapes getShapesFromPage()
        {
            Visio.Shapes vsoShapes;
            vsoShapes = Application.ActiveDocument.Pages.get_ItemU(1).Shapes;
            return vsoShapes;

        }
        //TODO: Raise issue when there are no incoming and/or outgoing flows.
        public void gatewayValidator(Visio.Shapes shapes, Visio.Document document)
        {
            //Insert rule set
            Visio.ValidationRuleSet gatewayValidatorRuleSet = document.Validation.RuleSets.Add("Gateway Validation");
            gatewayValidatorRuleSet.Description = "Verify that the gateways are correctly used in the document.";

            Visio.ValidationRule customRule1 = gatewayValidatorRuleSet.Rules.Add("distinctFlows2sequenceFlow");
            customRule1.Category = "Gateway";
            customRule1.Description = "Wenn ungleiche Sequenzflüsse zusammengeführt werden, muss der ausgehende Sequenzfluss DangerFlow sein";

            Visio.ValidationRule customRule2 = gatewayValidatorRuleSet.Rules.Add("equalFlows2distinctFlow");
            customRule2.Category = "Gateway";
            customRule2.Description = "Eingehende Sequenzflüsse sind ungleich dem ausgehenden Sequenzfluss. Muss gleich sein.";

            Boolean mixedFlows = false;
            var incomingShapes = new List<Visio.Shape>();
            var outgoingShapes = new List<Visio.Shape>();

            foreach (Visio.Shape shape in shapes)
            {
                if(shape.Master.NameU == "Gateway")
                {
                    Array incoming1Dshapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming1D, "");
                    Array outgoing1Dshapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "");
                    foreach (Object element in incoming1Dshapes)
                    {
                        incomingShapes.Add(shapes.get_ItemFromID((int)element));
                    }
                    foreach (Object element in outgoing1Dshapes)
                    {
                        outgoingShapes.Add(shapes.get_ItemFromID((int)element));
                    }

                    String comparisonShape = incomingShapes.First().Name;
                    foreach (Visio.Shape current in incomingShapes)
                    {
                        if(current.Master.NameU != comparisonShape)
                        {   //Wenn ungleiche Sequenzflüsse zusammengeführt werden, müssen ausgehende Sequenzflüsse DangerFlows sein.
                            mixedFlows = true;
                            foreach(Visio.Shape x in outgoingShapes)
                            {
                                if(x.Master.Name != "DangerFlow")
                                {
                                    Visio.ValidationIssue vsoValidationIssue = customRule1.AddIssue(x.ContainingPage, x);
                                    vsoValidationIssue.Ignored = false;
                                }
                            }
                            break;                           
                        }
                    }
                    //Alle eingehenden Sequenzflüsse sind gleich, aber mindestens ein ausgehender Sequenzfluss ist ungleich.
                    if(mixedFlows == false)
                    {
                        foreach(Visio.Shape element in outgoingShapes)
                        {
                            if(element.Master.NameU != comparisonShape)
                            {
                                customRule2.AddIssue(element.ContainingPage, element);
                            }
                        }
                    }
                }
            }
        }

        public void inspectionValidator(Visio.Shapes shapes, Visio.Document document)
        {
            // TODO: Add issue for missing sequence flow
            //Insert rule set
            Visio.ValidationRuleSet inspectionValidatorRuleSet = document.Validation.RuleSets.Add("Inspection Validation");
            inspectionValidatorRuleSet.Description = "Verify that the Inspection-Shapes are correctly used in the document.";

            Visio.ValidationRule customRule3 = inspectionValidatorRuleSet.Rules.Add("glued2DshapesMissing");
            customRule3.Category = "inspection-shape";
            customRule3.Description = "Kein 2D-Shape an das Inspection-Shape geklebt.";

            Visio.ValidationRule customRule4 = inspectionValidatorRuleSet.Rules.Add("gluedViolationEvent");
            customRule4.Category = "inspection-shape";
            customRule4.Description = "Kein Violation-Event an das Inspections-Shape geklebt.";

            var gluedShapesIDs = new List<int>();
            int count = 0;
            foreach(Visio.Shape shape in shapes)
            {
                if(shape.Master.NameU == "Inspection")
                {
                    Array glued2Dshapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll2D, ""); 
                    foreach(Object element in glued2Dshapes){
                        gluedShapesIDs.Add((int)element);
                    }
                    if (!gluedShapesIDs.Any())  
                    {
                        customRule3.AddIssue(shape.ContainingPage, shape); //Issue Handling    Keine glued 2D Shapes vorhanden
                        break;
                    }
                    else
                    {
                        foreach(int id in gluedShapesIDs)
                        {
                            if (shapes.get_ItemFromID(id).Master.Name == "Violation")
                            {
                                count++;
                            }
                        }
                        if (count== 0)
                        {
                            customRule4.AddIssue(shape.ContainingPage, shape); //Issue Handling    Keine "Violation"-Shape an das Inspection-Shape geklebt.
                        }
                    }
                }
            }
        }

        public void violationValidator(Visio.Shapes shapes, Visio.Document document)
        {
            var outgoingShapes = new List<Visio.Shape>();
            foreach(Visio.Shape shape in shapes)
            {
                if(shape.Name == "Violation")
                {
                    Array outgoing1Dshapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "");
                    foreach (Object element in outgoing1Dshapes)
                    {
                        outgoingShapes.Add(shapes.get_ItemFromID((int)element));
                    }
                    foreach(Visio.Shape element in outgoingShapes)
                    {
                        if(element.Master.Name != "DangerFlow")
                        {
                            //Issue Handling
                        }
                    }
                }
                
            }
        }

        public void surveillanceValidator(Visio.Shapes shapes, Visio.Document document)
        {
            var surveillanceShapes = new List<String>();
            surveillanceShapes.Add("SecurityGuard");
            surveillanceShapes.Add("CCTV");
            surveillanceShapes.Add("PerimeterBarrier");
            surveillanceShapes.Add("AlarmSystem");
            foreach (Visio.Shape shape in shapes)
            {
                if(surveillanceShapes.Contains(shape.Name))
                {
                    if(shape.ContainingShape.Name == "Group")       //If the Shape object is the member of a group, the ContainingShape property returns that group.
                    {
                        //Prüfe, ob die Gruppe einen ausgehenden Message flow hat
                        var outgoingFlows = new List<Visio.Shape>();
                        outgoingFlows = getOutgoingShapes(shape);
                        if(!outgoingFlows.Exists(x => x.Name == "MessageFlow")){
                            //Issue Handling    Kein Message Flow an Surveillance-Gruppe
                        }
                        
                    }
                    else if (shape.ContainingShape.Name == "Pool")
                    {
                        //
                    }
                    else
                    {
                        //Issue Handling    Surveillance Shape als top-level shape verwendet
                    }

                }
            }
        }

        public List<Visio.Shape> getIncomingShapes(Visio.Shape currentShape)
        {
        var shapes = new List<Visio.Shape>();
        Visio.Connects shapeFromConnections = currentShape.FromConnects;
        foreach(Visio.Connect connection in shapeFromConnections)
            {
            shapes.Add(connection.FromSheet);  // https://docs.microsoft.com/de-de/office/vba/api/visio.connects
            }
        return shapes;
        }
        public List<Visio.Shape> getOutgoingShapes(Visio.Shape currentShape)
        {
            var shapes = new List<Visio.Shape>();
            Visio.Connects shapeConnections = currentShape.Connects;
            foreach (Visio.Connect connection in shapeConnections)
            {
                shapes.Add(connection.ToSheet);  // https://docs.microsoft.com/de-de/office/vba/api/visio.connects
            }
            return shapes;
        }

        public String[] getShapeNames(Visio.Shapes shapes)         //https://docs.microsoft.com/de-de/office/vba/api/visio.shapes.item
        {
            Visio.Shapes vsoShapes = shapes;
            int shapeCount = vsoShapes.Count;
            string[] names = new string[shapeCount];
            if (shapeCount > 0)
            {
                for (int i = 1; i < shapeCount; i++)
                {
                    names[i - 1] = vsoShapes.get_ItemU(i).Name;
                }
            }
            return names;
        }




        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
