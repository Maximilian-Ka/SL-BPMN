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
            

        List<Visio.Shape> IssueList;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string docPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) + @"\test\myDrawing.vsdx";
            Visio.Document doc = this.Application.Documents.Open(docPath);
            Visio.Page page = doc.Pages.get_ItemU(1);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        public Visio.Shapes getShapesFromPage()
        {
            Visio.Shapes vsoShapes;
            vsoShapes = Application.ActiveDocument.Pages.get_ItemU(1).Shapes;
            return vsoShapes;

        }


        //geht davon aus, dass sequence flows oder danger flows in das gateway führen
        public void gatewayValidator(Visio.Shapes shapes)
        {
            foreach(Visio.Shape shape in shapes)
            {
                if(shape.Name == "Gateway")
                {
                    List<Visio.Shape>flowShapes = getIncomingShapes(shape);
                    String comparisonShape = flowShapes.ElementAt(0).Name;
                    foreach(Visio.Shape flowShape in flowShapes)
                    {
                        if(flowShape.Name != comparisonShape)
                        {   //Wenn ungleiche Sequenzflüsse zusammengeführt werden, muss der ausgehende Sequenzfluss DangerFlow sein.
                            if (shape.Connects.get_Item16(0).ToSheet.Name != "DangerFlow")
                            {
                                //Issue Handling
                            }
                        }
                    }
                    //Alle eingehenden Sequenzflüsse sind gleich, aber der ausgehende Sequenzfluss ist ungleich.
                    if(comparisonShape != shape.Connects.get_Item16(0).ToSheet.Name)
                    {
                        //Issue Handling
                    }
                }
            }
        }

        public void inspectionValidator(Visio.Shapes shapes)
        {
            var gluedShapesIDs = new List<int>();
            int count = 0;
            foreach(Visio.Shape shape in shapes)
            {
                if(shape.Name == "Inspektion")
                {
                    Array glued2dShapes = shape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesIncoming2D, "");    // If the source object is a 2D shape, return the 2D shapes that are glued to this shape.
                    foreach(Object element in glued2dShapes){
                        gluedShapesIDs.Add((int)element);
                    }
                    if (!gluedShapesIDs.Any())  
                    {
                        //Issue Handling    Keine glued 2D Shapes vorhanden
                    }
                    else
                    {
                        foreach(int ID in gluedShapesIDs)
                        {
                            if (shapes.get_ItemFromID(ID).Name == "Violation")
                            {
                                count++;
                            }
                        }
                        if (count== 0)
                        {
                            //Issue Handling    Keine "Violation"-Shape an das Inspection-Shape geklebt.
                        }
                    }
                }
            }
        }

        public void violationValidator(Visio.Shapes shapes)
        {
            var outgoingShapes = new List<Visio.Shape>();
            foreach(Visio.Shape shape in shapes)
            {
                if(shape.Name == "Violation")
                {
                    outgoingShapes = getOutgoingShapes(shape);
                    foreach(Visio.Shape element in outgoingShapes)
                    {
                        if(element.Name != "DangerFlow")
                        {
                            //Issue Handling
                        }
                    }
                }
                
            }
        }

        public void surveillanceValidator(Visio.Shapes shapes)
        {
            var surveillanceShapes = new List<String>();
            surveillanceShapes.Add("SecurityGuard");
            surveillanceShapes.Add("CCTV");
            surveillanceShapes.Add("");
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

        //public long[] getIncomingShapes1(Visio.Shape currentShape)       // https://docs.microsoft.com/de-de/office/vba/api/visio.shape.connectedshapes
        //{
        //    long[] lngShapeIDs;
        //    int intCount;
        //    Visio.VisConnectedShapesFlags test = Visio.VisConnectedShapesFlags.visConnectedShapesIncomingNodes;
        //long[] test1 = currentShape.ConnectedShapes(Visio.VisConnectedShapesFlags.visConnectedShapesAllNodes, "") as long[];
        //    lngShapeIDs = Array.ConvertAll(currentShape.ConnectedShapes(Visio.VisConnectedShapesFlags.visConnectedShapesAllNodes, ""), item => (long)item);  //.OfType<object>().Select(o => o.ToString()).ToArray();
        //    currentShape.ConnectedShapes(Visio.VisConnectedShapesFlags.visConnectedShapesAllNodes, "").CopyTo(lngShapeIDs, 0);
        //    return lngShapeIDs;



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
