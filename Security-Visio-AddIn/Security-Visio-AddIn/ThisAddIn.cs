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
        // TODO: Öffnen von Document nicht hardcoden.
        // TODO: Weitere Ausnahmen behandeln.
        // TODO: Issue Handling implementieren
        // TODO: Von Validator zu Validator kann die übergebene Liste an Shapes gekürzt werden, damit Shapes nicht immer wieder überprüft werden.
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Visio.Document doc;
            if (Application.Documents.Count > 0)
            {
                doc = Application.ActiveDocument;
            }
            else
            {
                string docPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) + @"\test\myDrawing.vsdx";
                doc = this.Application.Documents.Open(docPath);
            }
            Visio.Shapes vsoShapes = getShapesFromPage();
            Visio.ValidationRuleSet gatewayValidatorRuleSet = doc.Validation.RuleSets.Add("Gateway Validation");
            gatewayValidatorRuleSet.Description = "Verify that the gateways are correctly used in the document.";
            Visio.ValidationRuleSet surveillanceValidatorRuleSet = doc.Validation.RuleSets.Add("Surveillance Validation");
            surveillanceValidatorRuleSet.Description = "Verify that the Surveillance elements are correctly used in the document.";
            Visio.ValidationRuleSet inspectionValidatorRuleSet = doc.Validation.RuleSets.Add("Inspection Validation");
            inspectionValidatorRuleSet.Description = "Verify that the Inspection-Shapes are correctly used in the document.";
            Visio.ValidationRuleSet violationValidatorRuleSet = doc.Validation.RuleSets.Add("Violation Validation");
            violationValidatorRuleSet.Description = "Verify that the Violation events are correctly used in the document.";
            Visio.ValidationRuleSet ciaValidatorRuleSet = doc.Validation.RuleSets.Add("CIA Validation");
            ciaValidatorRuleSet.Description = "Verify that the CIA elements are correctly used in the document.";
            //When Microsoft Visio performs validation, it fires a RuleSetValidated event for every rule set that it processes, even if a rule set is empty.
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
            if (RuleSet.Name == "CIA Validation")
            {
                CIAValidator(getShapesFromPage(), getActiveDocument());
                return;
            }
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


        public void gatewayValidator(Visio.Shapes shapes, Visio.Document document)
        {
            // TODO: Issue Handling
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
                    if(incoming1Dshapes.Length == 0)
                    {
                        //Issue Handling    Keine eingehenden Flows
                        break;
                    }
                    if(outgoing1Dshapes.Length == 0)
                    {
                        //Issue Handling    Keine ausgehenden Flows
                        break;
                    }
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
            //Rule set
            Visio.ValidationRuleSet inspectionValidatorRuleSet = document.Validation.RuleSets.Add("Inspection Validation");
            inspectionValidatorRuleSet.Description = "Verify that the Inspection-Shapes are correctly used in the document.";

            Visio.ValidationRule customRule1 = inspectionValidatorRuleSet.Rules.Add("missingSequenceFlow");
            customRule1.Category = "ispection-shape";
            customRule1.Description = "As each Inspection differentiates between secure and unsecure, it has to have a outgoing Sequence Flow to represent the secure path";
            
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
                        //Issue Handling: Missing Outgoing Sequence Flow for secure case distinction
                        customRule1.AddIssue(shape.ContainingPage, shape);
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
            //Ruleset 
            Visio.ValidationRuleSet violationValidatorRuleSet = document.Validation.RuleSets.Add("Violation Validation");
            violationValidatorRuleSet.Description = "Verify that the Violation events are correctly used in the document.";

            Visio.ValidationRule customRule1 = violationValidatorRuleSet.Rules.Add("noOutgoingDangerFlow");
            customRule1.Category = "Violation Event";
            customRule1.Description = "The outgoing flow of an Violation event has to be a DangerFlow";

            Visio.ValidationRule customRule2 = violationValidatorRuleSet.Rules.Add("noDangerFlow");
            customRule2.Category = "Violation Event";
            customRule2.Description = "A Violation event has to have a outgoing DangerFlow";

            //program logic
            var outgoingShapes = new List<Visio.Shape>();
            foreach(Visio.Shape shape in shapes)
            {
                if(shape.Master.Name == "Violation")
                {
                    Array outgoing1Dshapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "");
                    if(outgoing1Dshapes.Length == 0)
                    {
                        // Issue Handling   Violation Event muss einen ausgehenden DangerFlow haben. (shape weil event markiert werden soll)
                        customRule2.AddIssue(shape.ContainingPage, shape);
                        continue;
                    }
                    foreach (Object element in outgoing1Dshapes)
                    {
                        outgoingShapes.Add(shapes.get_ItemFromID((int)element));
                    }
                    foreach(Visio.Shape element in outgoingShapes)
                    {
                        if(element.Master.Name != "DangerFlow")
                        {
                            //Issue Handling: Outgoing Flow ist kein DangerFlow (element weil Flow markiert werden soll)
                            customRule1.AddIssue(shape.ContainingPage, element);
                        }
                    }
                }              
            }
        }


        public void surveillanceValidator(Visio.Shapes shapes, Visio.Document document)
        {
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
            //Issue Handling: Ruleset
            Visio.ValidationRuleSet ciaValidatorRuleSet = document.Validation.RuleSets.Add("CIA Validation");
            ciaValidatorRuleSet.Description = "Verify that the CIA elements are correctly used in the document.";
            
            Visio.ValidationRule customRule1 = ciaValidatorRuleSet.Rules.Add("availabilityNoOutMsgFlow");
            customRule1.Category = "CIA Elements";
            customRule1.Description = "The Availability element can only be used outside of Data-elements (Dataobject/Database/Message) if it is attached to an outgoing Message Flow";

             Visio.ValidationRule customRule2 = ciaValidatorRuleSet.Rules.Add("attachISshapeToDataElement");
            customRule2.Category = "CIA Elements";
            customRule2.Description = "Information Security elements can usually only be attached to Data-elements (Dataobject/Database/Message). Availability can additionally represent the Availability of a Message Flow";

            //programm logic
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
                                    //Issue Handling: Ausgehender Fluss muss Nachrichtenfluss sein
                                    customRule1.AddIssue(shape.ContainingPage, shape);
                                }
                            }
                            if (!outgoingFlows.Any())
                            {
                                //Issue Handling: Braucht ausgehenden Nachrichtenfluss
                                customRule1.AddIssue(shape.ContainingPage, shape);
                            }
                            return;
                        }
                        

                    }
                    var gluedShapes = new List<Visio.Shape>();
                    Array allGluedShapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll2D, "");
                    if(allGluedShapes.Length == 0)
                    {
                        //Issue Handling    Data-Security Element muss an ein Data-Shape geklebt werden.
                    }
                    foreach (Object element in allGluedShapes)
                    {
                        gluedShapes.Add(shapes.get_ItemFromID((int)element));
                    }
                    foreach(Visio.Shape element in gluedShapes)
                    {
                        if (!informationShapes.Contains(element.Master.Name))
                        {
                            //Issue Handing: Information Security Shape muss an ein Daten-Shapen geklebt werden.
                            customRule2.AddIssue(shape.ContainingPage, shape);
                        }
                    }
                }
            }
        }

        public void EntrypointValidator(Visio.Shapes shapes, Visio.Document document)
        {
            //Issue Handling (temp)
            Visio.ValidationRuleSet entryValidatorRuleSet = document.Validation.RuleSets.Add("EntryPoint Validation");
            entryValidatorRuleSet.Description = "Verify that the CIA elements are correctly used in the document.";

            Visio.ValidationRule customRule1 = entryValidatorRuleSet.Rules.Add("noOutFlow");
            customRule1.Category = "EntryPoint";
            customRule1.Description = "An EntryPoint needs to have either an outgoing Sequence Flow or an outgoing DangerFlow";

            Visio.ValidationRule customRule2 = entryValidatorRuleSet.Rules.Add("noInFlow");
            customRule2.Category = "EntryPoint";
            customRule2.Description = "An EntryPoint needs to have a incoming Sequence Flow or DangerFlow";

            Visio.ValidationRule customRule3 = entryValidatorRuleSet.Rules.Add("");
            customRule3.Category = "EntryPoint";
            customRule3.Description = "Whenever an EntryPoint stand before a secure zone (Group with PerimeterBarrier), it has to be preceded by an Identification task";

            Visio.ValidationRule customRule4 = entryValidatorRuleSet.Rules.Add("");
            customRule4.Category = "EntryPoint";
            customRule4.Description = "The an EntryPoint following Element, has to be inside a seperate zone (inside a Group object)";


            // Listen für auf den EntryPoint folgende Shapes
            var out1DShapeList = new List<Visio.Shape>();
            var out2DShapeList = new List<Visio.Shape>();
            // Listen für dem EntryPoint vorhergehenden Shapes
            var in1DShapeList = new List<Visio.Shape>();
            var in2DShapeList = new List<Visio.Shape>();
            // Liste für Container Objecte
            var containerShapes = new List<Visio.Shape>();
            var containerMembers = new List<Visio.Shape>();
            Boolean inGroup = false;

            //Programmlogik (haha "Logik" xD das würde ja vorraussetzten dass da alles logisch wäre du N00b)
            // Laufe über alle Shapes des Dokuments
            foreach(Visio.Shape shape in shapes)
            {
                if(shape.Master.Name == "EntryPoint") //Für jeden EntryPoint im Dokument
                {
                    //Es wird davon ausgegangen, dass EntryPoint immer mit nur einem outgoing Sequenzfluss (bzw. D.F.) verbunden ist
                    //zusätzlich kann ein EntryPoint aber  noch mit einem MessageFlow verbunden sein
                    //prüfe alle outgoing 1D shapes (Pfeile: Sequenzfluss, DangerFlow, MSgFlow)
                    Array out1DArray = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "");
                    if(out1DArray.Length == 0)
                    {
                        customRule1.AddIssue(shape.ContainingPage, shape);
                    }
                    //Array in Liste mit den out 1D shapes casten
                    foreach (Object element in out1DArray)
                    {
                        out1DShapeList.Add(shapes.get_ItemFromID((int)element));
                    }
                    // Über alle outgoing Flows laufen
                    foreach(Visio.Shape outFlow in out1DShapeList)
                    {
                        if((outFlow.Master.Name == "DangerFlow") || (outFlow.Master.Name == "Sequenzfluss"))
                        {
                            Array out2DArray = outFlow.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing2D, ""); //an den Fluss gebundenen outgoing 2D shapes
                            //Es wird davon ausgegangen, dass nur ein outgoing 2D Shape an einen Flow gebunden ist
                            foreach(Object element in out2DArray)
                            {
                                out2DShapeList.Add(shapes.get_ItemFromID((int)element)); // in Liste mit den out 2D shapes casten
                            }
                            foreach(Visio.Shape out2DShape in out2DShapeList) //laufe über alle out 2D shapes (nur eins)
                            {
                                //ist das out2DShape Teil eines Gruppen-containers
                                Array containerIDs = out2DShape.MemberOfContainers; //Holt alle Container in denen das 2D Shape enthalten ist
                                foreach (Object element in containerIDs)
                                {
                                    containerShapes.Add(shapes.get_ItemFromID((int)element)); //Castet Container IDs in Liste mit Container Objekten
                                }
                                foreach(Visio.Shape container in containerShapes)
                                {
                                    //true wenn 2D shape Teil eines Gruppen container ist --> Issue Handling wenn 2D shape nicht Teil einer Gruppe ist (alle zugeordneten Container dürfen keine Gruppe sein)
                                    if(container.Master.NameU == "Group")
                                    {
                                        inGroup = true;
                                        //hole alle Shapes die Teil des Gruppen Containers sind
                                        Array memberShapes = container.ContainerProperties.GetMemberShapes(0); // Array "with all shape types and including items in nested containers"
                                        foreach(Object element in memberShapes) //kann nicht leer sein, da mind. out2DShape enthalten sein muss
                                        {
                                            containerMembers.Add(shapes.get_ItemFromID((int)element)); //in Liste mit allen Shapes auf dem Container casten
                                        }
                                        foreach(Visio.Shape memberShape in containerMembers) //prüfe alle Mitglieder des Gruppen containers
                                        {
                                            //wenn PreimeterBarrier enthalten -> Sicherheitsbereich -> Identifikation vor EntryPoint
                                            if (memberShape.Master.Name == "PerimeterBarrier") 
                                            {
                                                Array in1DArray = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming1D,"");
                                                if(in1DArray.Length == 0)
                                                {
                                                    // Issue Handling: EntryPoint benötigt einen Incoming Flow; Muss nach einem Identifikationselement stehen
                                                    customRule2.AddIssue(shape.ContainingPage, shape);
                                                }
                                                // In Liste mit in 1D Shapes casten
                                                foreach(Object element in in1DArray)
                                                {
                                                    in1DShapeList.Add(shapes.get_ItemFromID((int)element)); //es wird davon ausgegangen dass nur ein Element enthalten ist
                                                }
                                                //Laufe über alle incoming Flows (nur einer)
                                                foreach(Visio.Shape inFlow in in1DShapeList)
                                                {
                                                    Array in2DArray = inFlow.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming2D, "");
                                                    //In Liste mit allen 2D Shapes casten --> Liste enthält das 2D Shape welchses sich vor dem EntryPoint befindet
                                                    foreach(Object element in in2DArray)
                                                    {
                                                        in2DShapeList.Add(shapes.get_ItemFromID((int)element));
                                                    }
                                                    //prüfe vor vorangehendes Element eine Identifikation ist
                                                    foreach(Visio.Shape in2DShape in in2DShapeList)
                                                    {
                                                        if(in2DShape.Master.Name != "Identification")
                                                        {
                                                            // Issue Handling: Wenn EntryPoint vor einem Schutzbereich steht, muss vor dem EntryPoint eine Identifikation statt finden
                                                            customRule3.AddIssue(shape.ContainingPage, shape);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    
                                }
                                //Issue Handling: out 2D shape ist nicht Teil einer Gruppe
                                if(inGroup == false)
                                {
                                    // Issue Handling: out 2D shape ist nicht Teil einer Gruppe
                                    customRule4.AddIssue(out2DShape.ContainingPage, out2DShape);
                                }
                            }
                        }
                    }
                }
            }
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
