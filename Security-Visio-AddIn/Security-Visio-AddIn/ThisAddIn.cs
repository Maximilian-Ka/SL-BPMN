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
            insertRuleSets(doc);
            //Reihenfolge an Regeln nicht verändern!
            Visio.Shapes vsoShapes = getShapesFromPage();
 
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
                gatewayValidator(getShapesFromPage(), getActiveDocument(), RuleSet);
                return;
            }
            if (RuleSet.Name == "Inspection Validation")
            {
                inspectionValidator(getShapesFromPage(), getActiveDocument(), RuleSet);
                return;
            }
            if (RuleSet.Name == "Violation Validation")
            {
                violationValidator(getShapesFromPage(), getActiveDocument(), RuleSet);
                return;
            }
            if (RuleSet.Name == "Surveillance Validation")
            {
                surveillanceValidator(getShapesFromPage(), getActiveDocument(), RuleSet);
                return;
            }
            if (RuleSet.Name == "CIA Validation")
            {
                CIAValidator(getShapesFromPage(), getActiveDocument(), RuleSet);
                return;
            }
            if (RuleSet.Name == "EntryPoint Validation")
            {
                EntrypointValidator(getShapesFromPage(), getActiveDocument(), RuleSet);
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


        public void gatewayValidator(Visio.Shapes shapes, Visio.Document document, Visio.ValidationRuleSet gatewayValidatorRuleSet)
        {

            foreach (Visio.Shape shape in shapes)
            {
                if(shape.Master.NameU == "Gateway")
                {
                    Boolean mixedFlows = false;
                    var incomingShapes = new List<Visio.Shape>();
                    var outgoingShapes = new List<Visio.Shape>();
                    Array incoming1Dshapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming1D, "");
                    Array outgoing1Dshapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "");
                    if(incoming1Dshapes.Length == 0)
                    {
                        //Issue Handling    Keine eingehenden Flows
                        gatewayValidatorRuleSet.Rules[3].AddIssue(shape.ContainingPage, shape);
                        continue;
                    }
                    if(outgoing1Dshapes.Length == 0)
                    {
                        //Issue Handling    Keine ausgehenden Flows
                        gatewayValidatorRuleSet.Rules[3].AddIssue(shape.ContainingPage, shape);
                        continue;
                    }
                    foreach (Object element in incoming1Dshapes)
                    {
                        incomingShapes.Add(shapes.get_ItemFromID((int)element));
                    }
                    foreach (Object element in outgoing1Dshapes)
                    {
                        outgoingShapes.Add(shapes.get_ItemFromID((int)element));
                    }

                    String comparisonShape = incomingShapes.First().Master.Name;
                    foreach (Visio.Shape current in incomingShapes)
                    {
                        if(current.Master.Name != comparisonShape)
                        {   //Wenn ungleiche Sequenzflüsse zusammengeführt werden, müssen ausgehende Sequenzflüsse DangerFlows sein.
                            mixedFlows = true;
                            foreach(Visio.Shape x in outgoingShapes)
                            {
                                if(x.Master.Name != "DangerFlow")
                                {
                                    //customRule1.AddIssue(x.ContainingPage, x);
                                    gatewayValidatorRuleSet.Rules[1].AddIssue(shape.ContainingPage, shape);
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
                            if(element.Master.Name != comparisonShape)
                            {
                                //customRule2.AddIssue(element.ContainingPage, element);
                                gatewayValidatorRuleSet.Rules[2].AddIssue(shape.ContainingPage, shape);
                            }
                        }
                    }
                }
            }
        }

        public void inspectionValidator(Visio.Shapes shapes, Visio.Document document, Visio.ValidationRuleSet inspectionValidatorRuleSet)
        {

            int count = 0;
            foreach(Visio.Shape shape in shapes)
            {
                if(shape.Master.Name == "Inspection")
                {
                    var gluedShapesIDs = new List<int>();
                    if (shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "").Length != 1)
                    {
                        //Issue Handling: Missing Outgoing Sequence Flow for secure case distinction
                        inspectionValidatorRuleSet.Rules[1].AddIssue(shape.ContainingPage, shape);
                    }
                    Array glued2Dshapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll2D, ""); 
                    foreach(Object element in glued2Dshapes){
                        gluedShapesIDs.Add((int)element);
                    }
                    if (!gluedShapesIDs.Any())  
                    {
                        inspectionValidatorRuleSet.Rules[2].AddIssue(shape.ContainingPage, shape); // Issue Handling: Keine glued 2D Shapes vorhanden
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
                            inspectionValidatorRuleSet.Rules[3].AddIssue(shape.ContainingPage, shape); // Issue Handling: Keine "Violation"-Shape an das Inspection-Shape geklebt.
                        }
                    }
                }
            }
        }

        public void violationValidator(Visio.Shapes shapes, Visio.Document document, Visio.ValidationRuleSet violationValidatorRuleSet)
        {

            foreach(Visio.Shape shape in shapes)
            {
                if(shape.Master.Name == "Violation")
                {
                    var outgoingShapes = new List<Visio.Shape>();
                    Array outgoing1Dshapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "");
                    if(outgoing1Dshapes.Length == 0)
                    {
                        // Issue Handling: Violation Event muss einen ausgehenden DangerFlow haben. (shape weil event markiert werden soll)
                        //customRule2.AddIssue(shape.ContainingPage, shape);
                        violationValidatorRuleSet.Rules.get_ItemFromID(2).AddIssue(shape.ContainingPage, shape);
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
                            //customRule1.AddIssue(shape.ContainingPage, element);
                            violationValidatorRuleSet.Rules.get_ItemFromID(1).AddIssue(shape.ContainingPage, element);
                        }
                    }
                }              
            }
        }

        public void surveillanceValidator(Visio.Shapes shapes, Visio.Document document, Visio.ValidationRuleSet surveillanceValidatorRuleSet)
        {

            var surveillanceShapes = new List<String>();
            surveillanceShapes.Add("SecurityGuard");
            surveillanceShapes.Add("CCTV");
            surveillanceShapes.Add("AlarmSystem");
            surveillanceShapes.Add("SecurityGuardTask.54"); //kp warum der Master so heißt
            Boolean inGroup = false;

            foreach (Visio.Shape shape in shapes)
            {
                if (surveillanceShapes.Contains(shape.Master.Name))
                {
                    var containerShapes = new List<Visio.Shape>();
                    //Prüft ob es sich um die S.G._Task handelt
                    if (shape.Master.Name == "SecurityGuardTask.54")
                    {
                        // holt alle Flows von dem S.G._Task
                        var outgoingShapes = new List<Visio.Shape>();
                        Array outgoing1Dshapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "");
                        foreach (Object element in outgoing1Dshapes)
                        {
                            outgoingShapes.Add(shapes.get_ItemFromID((int)element));
                        }
                        Boolean hasMsgFlow = false;
                        foreach(Visio.Shape outFlow in outgoingShapes)
                        {
                            if(outFlow.Master.Name == "Nachrichtenfluss")
                            {
                                hasMsgFlow = true;
                            }
                        }
                        if(hasMsgFlow == false)
                        {
                            // Issue Handling: No outgoing Message Flow bei der SecurityGuardTask (Regel 4 [13])
                            surveillanceValidatorRuleSet.Rules[4].AddIssue(shape.ContainingPage, shape);
                        }
                        continue; //shape von weiterer Bearbeitung ausschließen
                    }
                    //Prüft ob dem Shape ein Container zugeordnet ist, wenn nicht: Verstoß gegen Modellierungsregel 1
                    if(shape.MemberOfContainers.Length == 0){
                        //customRule1.AddIssue(shape.ContainingPage, shape);
                        surveillanceValidatorRuleSet.Rules[1].AddIssue(shape.ContainingPage, shape);
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
                                Array outgoing1Dshapes = container.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "");
                                foreach (Object element in outgoing1Dshapes)
                                {
                                    outgoingShapes.Add(shapes.get_ItemFromID((int)element));
                                }
                                if (!outgoingShapes.Any())
                                {
                                    // Issue Handling: Kein outgoing Message Flow an dem überwachten Group-Shape
                                    //customRule2.AddIssue(shape.ContainingPage, shape);
                                    surveillanceValidatorRuleSet.Rules[2].AddIssue(shape.ContainingPage, shape);
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
                                    var cffOfCont = new List<Visio.Shape>();
                                    Array cont = container.MemberOfContainers;
                                    // erstellt eine Liste mit den CFF containern des Swimlane List containers (hoffentlich immer nur einer) ->Funktioniert das?
                                    // die Konnektoren hängen nur an den CFF-Containern: der CFF-Container beinhaltet eine "Swimlane List" und eine "Phasen List" (Container Elemente "Pool/Lane" dafür nicht nutzbar warum auch immer)
                                    foreach(Object element in cont)
                                    {
                                        if(shapes.get_ItemFromID((int)element).Master.Name == "CFF-Container")
                                        {
                                            cffOfCont.Add(shapes.get_ItemFromID((int)element));
                                        }
                                    }
                                    //Überspringt wahrscheinlich die schleife --> If abfrage??
                                    foreach (Visio.Shape cff in cffOfCont)
                                    {
                                        var outgoingShapes = new List<Visio.Shape>();
                                        Array outgoing1Dshapes = cff.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "");
                                        foreach (Object element in outgoing1Dshapes)
                                        {
                                            outgoingShapes.Add(shapes.get_ItemFromID((int)element));
                                        }
                                        //hat der cff container outgoing 1D shapes?
                                        if (!outgoingShapes.Any())
                                        {
                                            // Issue Handling: Kein outgoing Message Flow an dem überwachten Lane-Shape
                                            //customRule3.AddIssue(shape.ContainingPage, shape);
                                            surveillanceValidatorRuleSet.Rules[3].AddIssue(shape.ContainingPage, shape);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        public void CIAValidator(Visio.Shapes shapes, Visio.Document document, Visio.ValidationRuleSet ciaValidatorRuleSet)
        {

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
                                    //customRule1.AddIssue(shape.ContainingPage, shape);
                                    ciaValidatorRuleSet.Rules[1].AddIssue(shape.ContainingPage, shape);
                                }
                            }
                            if (!outgoingFlows.Any())
                            {
                                //Issue Handling: Braucht ausgehenden Nachrichtenfluss
                                //customRule1.AddIssue(shape.ContainingPage, shape);
                                ciaValidatorRuleSet.Rules[1].AddIssue(shape.ContainingPage, shape);
                            }
                            return;
                        }
                    }
                    var gluedShapes = new List<Visio.Shape>();
                    Array allGluedShapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesAll2D, "");
                    if(allGluedShapes.Length == 0)
                    {
                        // Issue Handling: Data-Security Element muss an ein Data-Shape geklebt werden.
                        ciaValidatorRuleSet.Rules[2].AddIssue(shape.ContainingPage, shape);
                        continue;
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
                            //customRule2.AddIssue(shape.ContainingPage, shape);
                            ciaValidatorRuleSet.Rules[2].AddIssue(shape.ContainingPage, shape);
                        }
                    }
                }
            }
        }

        public void EntrypointValidator(Visio.Shapes shapes, Visio.Document document, Visio.ValidationRuleSet entrypointValidatorRuleSet)
        {

            // Laufe über alle Shapes des Dokuments
            foreach(Visio.Shape shape in shapes)
            {
                if(shape.Master.Name == "PerimeterBarrier")
                {
                    var containerShapes = new List<Visio.Shape>();
                    Array containerIDs = shape.MemberOfContainers; //Holt alle Container in denen das 2D Shape enthalten ist
                    foreach (Object element in containerIDs)
                    {
                        containerShapes.Add(shapes.get_ItemFromID((int)element)); //Castet Container IDs in Liste mit Container Objekten
                    }
                }
                if(shape.Master.Name == "EntryPoint") //Für jeden EntryPoint im Dokument
                {
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

                    //Es wird davon ausgegangen, dass EntryPoint immer mit nur einem outgoing Sequenzfluss (bzw. D.F.) verbunden ist
                    //zusätzlich kann ein EntryPoint aber  noch mit einem MessageFlow verbunden sein
                    //prüfe alle outgoing 1D shapes (Pfeile: Sequenzfluss, DangerFlow, MSgFlow)
                    Array out1DArray = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "");
                    if(out1DArray.Length == 0)
                    {
                        // Issue Handling: EntryPoint has to have an outgoing Flow
                        //customRule1.AddIssue(shape.ContainingPage, shape);
                        entrypointValidatorRuleSet.Rules[1].AddIssue(shape.ContainingPage, shape);
                        continue;
                    }
                    //Array in Liste mit den out 1D shapes casten
                    foreach (Object element in out1DArray)
                    {
                        out1DShapeList.Add(shapes.get_ItemFromID((int)element));
                    }
                    // Über alle outgoing Flows laufen
                    foreach(Visio.Shape outFlow in out1DShapeList)
                    {
                        if(outFlow.Master.Name == "DangerFlow" || outFlow.Master.Name == "Sequenzfluss")
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
                                                    //customRule2.AddIssue(shape.ContainingPage, shape);
                                                    entrypointValidatorRuleSet.Rules[2].AddIssue(shape.ContainingPage, shape);
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
                                                            //customRule3.AddIssue(shape.ContainingPage, shape);
                                                            entrypointValidatorRuleSet.Rules[3].AddIssue(shape.ContainingPage, shape);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                if(inGroup == false)
                                {
                                    // Issue Handling: out 2D shape ist nicht Teil einer Gruppe
                                    //customRule4.AddIssue(out2DShape.ContainingPage, out2DShape);
                                    entrypointValidatorRuleSet.Rules[4].AddIssue(shape.ContainingPage, shape);
                                }
                            }
                        }
                    }
                }
            }
        }

        public void insertRuleSets(Visio.Document doc)
        {
            //Ruleset für korrekte Nutzung von Gateways in Verbindung mit DangerFlows und Sequenzflüssen
            Visio.ValidationRuleSet gatewayValidatorRuleSet = doc.Validation.RuleSets.Add("Gateway Validation");
            gatewayValidatorRuleSet.Description = "Verify that the gateways are correctly used in the document.";
            Visio.ValidationRule customRule1 = gatewayValidatorRuleSet.Rules.Add("distinctFlows2sequenceFlow");
            customRule1.Category = "Gateway";
            customRule1.Description = "A DangerFlow and a regular SequenceFlow can only be combined in a gateway, if the resulting flow is a DangerFlow";
            Visio.ValidationRule customRule2 = gatewayValidatorRuleSet.Rules.Add("equalFlows2distinctFlow");
            customRule2.Category = "Gateway";
            customRule2.Description = "If incoming flows to a gateway are all of the same type, the outgoing flow must be of that same type";
            Visio.ValidationRule customRule3 = gatewayValidatorRuleSet.Rules.Add("noFlowsAttached");
            customRule3.Category = "Gateway";
            customRule3.Description = "A Gateway element has to have outcoming and incoming Flows";

            //Ruleset für die Elemente CCTV, Guard, AlarmSystem
            Visio.ValidationRuleSet surveillanceValidatorRuleSet = doc.Validation.RuleSets.Add("Surveillance Validation");
            surveillanceValidatorRuleSet.Description = "Verify that the Surveillance elements are correctly used in the document.";
            Visio.ValidationRule customRule10 = surveillanceValidatorRuleSet.Rules.Add("notAccociated");
            customRule10.Category = "Surveillance Element";
            customRule10.Description = "An Surveillance element has to be associated with a container object (Pool/Lane/Group)";
            Visio.ValidationRule customRule11 = surveillanceValidatorRuleSet.Rules.Add("noOutMsgGroup");
            customRule11.Category = "Surveillance Element";
            customRule11.Description = "The Group-Object associated with an Surveillance element has to have an outgoing MessageFlow";
            Visio.ValidationRule customRule12 = surveillanceValidatorRuleSet.Rules.Add("noOutMsgPool");
            customRule12.Category = "Surveillance Element";
            customRule12.Description = "The Pool/Lane-Object associated with an Surveillance element has to have an outgoing MessageFlow";
            Visio.ValidationRule customRule13 = surveillanceValidatorRuleSet.Rules.Add("noOutMsgSGT");
            customRule13.Category = "Surveillance Element";
            customRule13.Description = "A SecurityGuardTask task has to have an outgoing Message Flow";

            //Ruleset für das Inspektionselement
            Visio.ValidationRuleSet inspectionValidatorRuleSet = doc.Validation.RuleSets.Add("Inspection Validation");
            inspectionValidatorRuleSet.Description = "Verify that the Inspection-Shapes are correctly used in the document.";
            Visio.ValidationRule customRule20 = inspectionValidatorRuleSet.Rules.Add("missingSequenceFlow");
            customRule20.Category = "ispection-shape";
            customRule20.Description = "As each Inspection differentiates between secure and unsecure, it has to have a outgoing Sequence Flow to represent the secure path";
            Visio.ValidationRule customRule21 = inspectionValidatorRuleSet.Rules.Add("glued2DshapesMissing");
            customRule21.Category = "inspection-shape";
            customRule21.Description = "As each Inspection differentiates between secure and unsecure, a Violation event needs to be glued to a Inspection task, to represent the start of a DangerFlow";
            Visio.ValidationRule customRule22 = inspectionValidatorRuleSet.Rules.Add("gluedViolationEvent");
            customRule22.Category = "inspection-shape";
            customRule22.Description = "As each Inspection differentiates between secure and unsecure, a Violation event needs to be glued to a Inspection task, to represent the start of a DangerFlow";

            //Ruleset für Violation event
            Visio.ValidationRuleSet violationValidatorRuleSet = doc.Validation.RuleSets.Add("Violation Validation");
            violationValidatorRuleSet.Description = "Verify that the Violation events are correctly used in the document.";
            Visio.ValidationRule customRule30 = violationValidatorRuleSet.Rules.Add("noOutgoingDangerFlow");
            customRule30.Category = "Violation Event";
            customRule30.Description = "The outgoing flow of a Violation event has to be a DangerFlow";
            Visio.ValidationRule customRule31 = violationValidatorRuleSet.Rules.Add("noDangerFlow");
            customRule31.Category = "Violation Event";
            customRule31.Description = "A Violation event has to have a outgoing DangerFlow";

            //Ruleset für Elemente der Informationssicherheit
            Visio.ValidationRuleSet ciaValidatorRuleSet = doc.Validation.RuleSets.Add("CIA Validation");
            ciaValidatorRuleSet.Description = "Verify that the CIA elements are correctly used in the document.";
            Visio.ValidationRule customRule40 = ciaValidatorRuleSet.Rules.Add("availabilityNoOutMsgFlow");
            customRule40.Category = "CIA Elements";
            customRule40.Description = "The Availability element can only be used outside of Data-elements (Dataobject/Database/Message) if it is attached to an outgoing Message Flow";
            Visio.ValidationRule customRule41 = ciaValidatorRuleSet.Rules.Add("attachISshapeToDataElement");
            customRule41.Category = "CIA Elements";
            customRule41.Description = "Information Security elements can usually only be attached to Data-elements (Dataobject/Database/Message). Availability can additionally represent the Availability of a Message Flow";

            //Ruleset für EntryPoints
            Visio.ValidationRuleSet entryValidatorRuleSet = doc.Validation.RuleSets.Add("EntryPoint Validation");
            entryValidatorRuleSet.Description = "Verify that the CIA elements are correctly used in the document.";
            Visio.ValidationRule customRule50 = entryValidatorRuleSet.Rules.Add("noOutFlow");
            customRule50.Category = "EntryPoint";
            customRule50.Description = "An EntryPoint needs to have either an outgoing Sequence Flow or an outgoing DangerFlow";
            Visio.ValidationRule customRule51 = entryValidatorRuleSet.Rules.Add("noInFlow");
            customRule51.Category = "EntryPoint";
            customRule51.Description = "An EntryPoint needs to have a incoming Sequence Flow or DangerFlow";
            Visio.ValidationRule customRule52 = entryValidatorRuleSet.Rules.Add("noIdentification");
            customRule52.Category = "EntryPoint";
            customRule52.Description = "Whenever an EntryPoint stand before a secure zone (Group with PerimeterBarrier), it has to be preceded by an Identification task";
            Visio.ValidationRule customRule53 = entryValidatorRuleSet.Rules.Add("no seperate zone");
            customRule53.Category = "EntryPoint";
            customRule53.Description = "The an EntryPoint following element, has to be inside a seperate zone (inside a Group object)";
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
