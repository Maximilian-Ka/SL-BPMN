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
        public String[] getShapeNames(Visio.Shapes shapes)         //https://docs.microsoft.com/de-de/office/vba/api/visio.shapes.item
        {
            Visio.Shapes vsoShapes = shapes;
            int shapeCount = vsoShapes.Count;
            string[] names = new string[shapeCount];
            if(shapeCount > 0)
            {
                for(int i=1; i<shapeCount; i++)
                {
                    names[i - 1] = vsoShapes.get_ItemU(i).Name;
                }
            }
            return names;
        }

        public void gatewayValidator(Visio.Shapes shapes)
        {
            foreach(Visio.Shape shape in shapes)
            {
                if(shape.Name == "Gateway")
                {
                    getIncomingShapes(shape);
                }
            }
        }

        public List<Visio.Shape> getIncomingShapes(Visio.Shape currentShape)
        {
        var shapes = new List<Visio.Shape>();
        Visio.Connects shapeFromConnects = currentShape.FromConnects;
        foreach(Visio.Connect connect in shapeFromConnects)
            {
            shapes.Add(connect.FromSheet);
            }
        return shapes;
        }

        public long[] getIncomingShapes1(Visio.Shape currentShape)       // https://docs.microsoft.com/de-de/office/vba/api/visio.shape.connectedshapes
        {
            long[] lngShapeIDs;
            int intCount;
            Visio.VisConnectedShapesFlags test = Visio.VisConnectedShapesFlags.visConnectedShapesIncomingNodes;
            long[] test1 = currentShape.ConnectedShapes(Visio.VisConnectedShapesFlags.visConnectedShapesAllNodes, "");
            //lngShapeIDs = Array.ConvertAll(currentShape.ConnectedShapes(Visio.VisConnectedShapesFlags.visConnectedShapesAllNodes, ""), item => (long)item);  //.OfType<object>().Select(o => o.ToString()).ToArray();
            currentShape.ConnectedShapes(Visio.VisConnectedShapesFlags.visConnectedShapesAllNodes, "").CopyTo(lngShapeIDs, 0);
            return lngShapeIDs;
        }
        //Retrieve every shape object that is connected(glued) to currentShape      Connect-object
        public boolean dangerflowInGateway(Visio.Shape currentShape)   
        {
            Visio.Connects connections = currentShape.FromConnects;
            for(int i=1; i<connections.Count; i++)
            {
                if (connections.Item(i).ObjectType==visObjTypeColor)
                {
                
                }
            }
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
