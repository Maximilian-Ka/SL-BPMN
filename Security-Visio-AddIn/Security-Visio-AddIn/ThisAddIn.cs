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
            Visio.ValidationRuleSet surveillanceValidatorRuleSet = doc.Validation.RuleSets.Add("Surveillance Validation");
            surveillanceValidatorRuleSet.Description = "Verify that the Surveillance elements are correctly used in the document.";
            Visio.ValidationRuleSet inspectionValidatorRuleSet = doc.Validation.RuleSets.Add("Inspection Validation");
            inspectionValidatorRuleSet.Description = "Verify that the Inspection-Shapes are correctly used in the document.";
            Application.RuleSetValidated += new Visio.EApplication_RuleSetValidatedEventHandler(HandleRuleSetValidatedEvent);
            //var incomingShapes = new List<String>();
            //foreach(Visio.Shape test in vsoShapes)
            //{
            //    incomingShapes.Add(test.Master.NameU);
            //}
            //foreach(Visio.Shape test in vsoShapes)
            //{
             //   test.ReplaceShape(doc.Masters.get_ItemU("DangerFlow"), 0);
            //}
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void HandleRuleSetValidatedEvent(Visio.ValidationRuleSet RuleSet)
        {
            //gatewayValidator(getShapesFromPage(), getActiveDocument());
            if (RuleSet.Name == "Gateway Validation")
            {
                gatewayValidator(getShapesFromPage(), getActiveDocument());
                return;
            }
            if (RuleSet.Name == "Inspection Validation")
            {
                inspectionValidator(getShapesFromPage(), getActiveDocument());
                return;
            }
            if (RuleSet.Name == "Violation Validation")
            {
                violationValidator(getShapesFromPage(), getActiveDocument());
                return;
            }
            if (RuleSet.Name == "Surveillance Validation")
            {
                surveillanceValidator(getShapesFromPage(), getActiveDocument());
                return;
            }
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
            customRule1.Description = "A DangerFlow and a regular SequenceFlow can only be combined in a gateway, if the resulting flow is a DangerFlow";

            Visio.ValidationRule customRule2 = gatewayValidatorRuleSet.Rules.Add("equalFlows2distinctFlow");
            customRule2.Category = "Gateway";
            customRule2.Description = "If incoming flows to a gateway are all of the same type, the outgoing flow must be of that same type";

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
                                    customRule1.AddIssue(x.ContainingPage, x);
                                    
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
            //Rule set
            Visio.ValidationRuleSet inspectionValidatorRuleSet = document.Validation.RuleSets.Add("Inspection Validation");
            inspectionValidatorRuleSet.Description = "Verify that the Inspection-Shapes are correctly used in the document.";

            Visio.ValidationRule customRule3 = inspectionValidatorRuleSet.Rules.Add("glued2DshapesMissing");
            customRule3.Category = "inspection-shape";
            customRule3.Description = "As each Inspection differentiates between secure and unsecure, a Violation event needs to be glued to a Inspection task, to represent the start of a DangerFlow";

            Visio.ValidationRule customRule4 = inspectionValidatorRuleSet.Rules.Add("gluedViolationEvent");
            customRule4.Category = "inspection-shape";
            customRule4.Description = "As each Inspection differentiates between secure and unsecure, a Violation event needs to be glued to a Inspection task, to represent the start of a DangerFlow";

            var gluedShapesIDs = new List<int>();
            int count = 0;
            foreach(Visio.Shape shape in shapes)
            {
                if(shape.Master.Name == "Inspection")
                {
                    if(shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "").Length != 1)
                    {
                        //Issue Handling    Muss genau einen outgoing Sequenzfluss haben.
                    }
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
            // TODO: Issue Handling
            //Ruleset 
            Visio.ValidationRuleSet violationValidatorRuleSet = document.Validation.RuleSets.Add("Violation Validation");
            violationValidatorRuleSet.Description = "Verify that the Violation events are correctly used in the document.";

            Visio.ValidationRule customRule1 = violationValidatorRuleSet.Rules.Add("noOutgoingDangerFlow");
            customRule1.Category = "Violation Event";
            customRule1.Description = "The outgoing flow of an Violation event has to be a DangerFlow";

            var outgoingShapes = new List<Visio.Shape>();
            foreach(Visio.Shape shape in shapes)
            {
                if(shape.Master.Name == "Violation")
                {
                    Array outgoing1Dshapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "");
                    if(outgoing1Dshapes.Length == 0)
                    {
                        // Issue Handling   Violation Event muss einen ausgehenden DangerFlow haben.
                        break;
                    }
                    foreach (Object element in outgoing1Dshapes)
                    {
                        outgoingShapes.Add(shapes.get_ItemFromID((int)element));
                    }
                    foreach(Visio.Shape element in outgoingShapes)
                    {
                        if(element.Master.Name != "DangerFlow")
                        {
                            customRule1.AddIssue(shape.ContainingPage, element);
                        }
                    }
                }              
            }
        }


        public void surveillanceValidator(Visio.Shapes shapes, Visio.Document document)
        {
            // TODO: Issue handling
            //Ruleset
            Visio.ValidationRuleSet surveillanceValidatorRuleSet = document.Validation.RuleSets.Add("Surveillance Validation");
            surveillanceValidatorRuleSet.Description = "Verify that the Surveillance elements are correctly used in the document.";

            Visio.ValidationRule customRule1 = surveillanceValidatorRuleSet.Rules.Add("notAccociated");
            customRule1.Category = "Surveillance Element";
            customRule1.Description = "An Surveillance element has to be associated with a container object (Pool/Lane/Group)";

            Visio.ValidationRule customRule2 = surveillanceValidatorRuleSet.Rules.Add("noOutMsgGroup");
            customRule2.Category = "Surveillance Element";
            customRule2.Description = "The Group-Object associated with an Surveillance element has to have an outgoing MessageFlow";

            Visio.ValidationRule customRule3 = surveillanceValidatorRuleSet.Rules.Add("noOutMsgPool");
            customRule3.Category = "Surveillance Element";
            customRule3.Description = "The Pool/Lane-Object associated with an Surveillance element has to have an outgoing MessageFlow";

            //Validator
            var surveillanceShapes = new List<String>();
            var containerShapes = new List<Visio.Shape>();
            surveillanceShapes.Add("SecurityGuard");
            surveillanceShapes.Add("CCTV");
            surveillanceShapes.Add("PerimeterBarrier");
            surveillanceShapes.Add("AlarmSystem");
            Boolean inGroup = false;
            foreach (Visio.Shape shape in shapes)
            {
                if(surveillanceShapes.Contains(shape.Master.Name))
                {
                    //TODO Prüfe, ob outgoing1Dflows MessageFlows sind.
                    //Prüft ob dem Shape ein Container zugeordnet ist, wenn nicht: Verstoß gegen Modellierungsregel 1
                    if(shape.MemberOfContainers == null){
                        customRule1.AddIssue(shape.ContainingPage, shape);
                    }
                    else{
                        //Holt alle Container Objekte in denen das aktuelle Shape liegt; Fügt sie der Liste containerIDs hinzu
                        Array containerIDs = shape.MemberOfContainers;
                        foreach(Object element in containerIDs){
                            containerShapes.Add(shapes.get_ItemFromID((int)element));
                        }

                        foreach(Visio.Shape container in containerShapes){
                            if(container.Master.NameU == "Group"){
                                var outgoingShapes = new List<Visio.Shape>();
                                Array outgoing1Dshapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "");
                                foreach (Object element in outgoing1Dshapes)
                                {
                                    outgoingShapes.Add(shapes.get_ItemFromID((int)element));
                                }
                                if (!outgoingShapes.Any())
                                {
                                    //Issue  Kein outgoing Message Flow an dem überwachten Group-Shape
                                    customRule2.AddIssue(shape.ContainingPage, shape);
                                }
                                inGroup = true;
                                break;
                            }
                        }
                        //Surveillance Shape in einer Lane/in einem Pool.
                        if (inGroup == false)
                        {
                            foreach (Visio.Shape container in containerShapes)
                            {
                                if (container.Master.NameU == "Swimlane List")
                                {
                                    var outgoingShapes = new List<Visio.Shape>();
                                    Array outgoing1Dshapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "");
                                    foreach (Object element in outgoing1Dshapes)
                                    {
                                        outgoingShapes.Add(shapes.get_ItemFromID((int)element));
                                    }
                                    if (!outgoingShapes.Any())
                                    {
                                        //Issue  Kein outgoing Message Flow an dem überwachten Lane-Shape
                                        customRule3.AddIssue(shape.ContainingPage, shape);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        public void CIAValidator(Visio.Shapes shapes, Visio.Document document)
        {
            // TODO: Issue handling
            var informationSecurityShapes = new List<String>();
            informationSecurityShapes.Add("Confidentiality");
            informationSecurityShapes.Add("Integrity");
            informationSecurityShapes.Add("Availability");
            var informationShapes = new List<String>();
            informationShapes.Add("Datenobjekt");
            informationShapes.Add("Nachricht");
            informationShapes.Add("Datenspeicher");
            foreach (Visio.Shape shape in shapes)
            {
                if (informationSecurityShapes.Contains(shape.Master.Name))
                {   
                    if(shape.Master.Name == "Availability")
                    {
                        Boolean gluedToNonDataObject = false;
                        var gluedShapes1 = new List<Visio.Shape>();
                        Array allGlued1Shapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll2D, "");
                        foreach (Object element in allGlued1Shapes)
                        {
                            gluedShapes1.Add(shapes.get_ItemFromID((int)element));
                        }
                        foreach (Visio.Shape element in gluedShapes1)
                        {
                            if (!informationShapes.Contains(element.Master.Name))
                            {
                                gluedToNonDataObject = true;
                            }
                        }
                        if (gluedToNonDataObject == true)
                        {
                            var outgoingFlows = new List<Visio.Shape>();
                            Array outgoing1D = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "");
                            foreach (Object element in outgoing1D)
                            {
                                outgoingFlows.Add(shapes.get_ItemFromID((int)element));
                            }
                            foreach(Visio.Shape x in outgoingFlows)
                            {
                                if(x.Master.Name != "Nachrichtenfluss")
                                {

                                    //Issue Handling    Ausgehender Fluss muss Nachrichtenfluss sein
                                }
                            }
                            if (!outgoingFlows.Any())
                            {
                                //Issue Handling    Braucht ausgehenden Nachrichtenfluss
                            }
                            return;
                        }
                        

                    }
                    var gluedShapes = new List<Visio.Shape>();
                    Array allGluedShapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll2D, "");
                    foreach (Object element in allGluedShapes)
                    {
                        gluedShapes.Add(shapes.get_ItemFromID((int)element));
                    }
                    foreach(Visio.Shape element in gluedShapes)
                    {
                        if (!informationShapes.Contains(element.Master.Name))
                        {
                            //Issue Handing     Information Security Shape muss an ein Daten-Shapen geklebt werden.
                        }
                    }
                }
            }
        }

        public void EntrypointValidator(Visio.Shapes shapes, Visio.Document document)
        {
            //  Shape.ContainerProperties.GetMemberState();
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
