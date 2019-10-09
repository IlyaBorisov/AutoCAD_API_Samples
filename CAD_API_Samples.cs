using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System.IO;
using sd = System.Diagnostics;
using System.Runtime.Serialization.Json;
using System.Runtime.Serialization;
using Autodesk.AutoCAD.Windows;

namespace ElectricPanels
{
  public class Node
    {

        public double power { get; set; }
        public double distance { get; set; }
        public double moment { get; set; }
    }

    public class BranchNode : Node
    {
        public Polyline poly { get; set; }
        public bool ischild { get; set; }
        public SortedDictionary<double, Node> nodes { get; set; }
    }
    public class Main : IExtensionApplication
    {
        public static string filepath = Autodesk.AutoCAD.ApplicationServices.Core.Application.GetSystemVariable("DWGPREFIX").ToString();
        public static string jsonfilename = filepath + Autodesk.AutoCAD.ApplicationServices.Core.Application.GetSystemVariable("DWGNAME").ToString().Replace("dwg", "json");
        public static double[] sechen = { 1.5, 2.5, 4, 6, 10, 16, 25, 35, 50, 70, 95, 120, 150, 185, 240 };
        public static int[] tokikabelei = { 21, 27, 36, 46, 63, 84, 112, 137, 167, 211, 261, 302, 346, 397, 472 };
        public static int[] switches = { 6, 10, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125, 160, 250, 400, 630, 800, 1000 };
        public static int[] rubilniki = { 20, 25, 32, 40, 63, 80, 100, 125 };
        public static int[] ShRNs = { 12, 24, 36, 48, 54, 72 };
        public static double lostmax = 2.4;
        
        public static List<T> GetObjectsFromDB<T>(Database db, string[] layers = null, string[] blocknames = null, int[] colors = null, bool hyperlinked = false) where T : class
        {
            var objs = new List<T>();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                foreach (ObjectId id in btr)
                {
                    Entity ent = (Entity)tr.GetObject(id, OpenMode.ForRead);
                    if (ent.GetType() == typeof(T) &&
                        (layers == null ? true : Array.IndexOf(layers, ent.Layer) >= 0) &&
                        (blocknames == null ? true : Array.IndexOf(blocknames, ((BlockReference)ent).Name) >= 0) &&
                        (colors == null ? true : Array.IndexOf(colors, ent.ColorIndex) >= 0) &&
                        (hyperlinked ? ent.Hyperlinks.Count > 0 : true))
                    {
                        T obj = ent as T;
                        objs.Add(obj);
                    }
                }
                tr.Commit();
            }
            return objs;
        }

        public static List<T> GetObjectsFromSelection<T>(Database db, Editor ed, string[] layers = null, string[] blocknames = null, int[] colors = null) where T : class
        {
            var objs = new List<T>();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var pso = new PromptSelectionOptions();
                pso.RejectObjectsOnLockedLayers = true;
                PromptSelectionResult psr = ed.GetSelection(pso);
                if (psr.Status == PromptStatus.OK)
                {
                    SelectionSet ss = psr.Value;
                    foreach (SelectedObject item in ss)
                    {
                        if (item != null)
                        {
                            Entity ent = (Entity)tr.GetObject(item.ObjectId, OpenMode.ForRead);
                            if (ent.GetType() == typeof(T) &&
                                (layers == null ? true : Array.IndexOf(layers, ent.Layer) >= 0) &&
                                (blocknames == null ? true : Array.IndexOf(blocknames, ((BlockReference)ent).Name) >= 0) &&
                                (colors == null ? true : Array.IndexOf(colors, ent.ColorIndex) >= 0))
                            {
                                T obj = ent as T;
                                objs.Add(obj);
                            }
                        }
                    }
                }
                tr.Commit();
            }
            return objs;
        }

        public static void DrawText(Database db, string content, Point position)
        {
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                var text = new DBText();
                text.SetDatabaseDefaults();                
                text.Height = 50;
                text.ColorIndex = 5;
                text.Position = new Point3d(position.X, position.Y, 0);                
                text.TextString = content;
                btr.AppendEntity(text);
                tr.AddNewlyCreatedDBObject(text, true);
                tr.Commit();
            }
        }
        
        public static double Moment(CableLine cable)
        {
            if (cable.polys.Count == 0 || cable.loads.Count == 0)
            {
                return 0;
            }

            var treelines = new List<BranchNode>();

            foreach (var poly in cable.polys)
            {
                var treeline = new BranchNode() { poly = poly, nodes = new SortedDictionary<double, Node>() };
                treelines.Add(treeline);
            }

            for (int i = 0; i < treelines.Count; i++)
            {
                for (int j = 0; j < treelines.Count; j++)
                {
                    if (treelines[i].poly != treelines[j].poly && (Distance(treelines[i].poly.GetClosestPointTo(treelines[j].poly.StartPoint, false), treelines[j].poly.StartPoint) < 10 || Distance(treelines[i].poly.GetClosestPointTo(treelines[j].poly.EndPoint, false), treelines[j].poly.EndPoint) < 10))
                    {
                        if (Distance(treelines[i].poly.GetClosestPointTo(treelines[j].poly.EndPoint, false), treelines[j].poly.EndPoint) < 10)
                            ReversePoly(treelines[j].poly);
                        treelines[i].nodes.Add(treelines[i].poly.GetDistAtPoint(treelines[i].poly.GetClosestPointTo(treelines[j].poly.StartPoint, false)), treelines[j]);
                        treelines[j].ischild = true;
                    }
                }
            }
            foreach (var treeline in treelines)
            {
                foreach (var load in cable.loads)
                {
                    if (Distance(treeline.poly.GetClosestPointTo(load.position, false), load.position) < 10)
                    {
                        treeline.nodes.Add(treeline.poly.GetDistAtPoint(treeline.poly.GetClosestPointTo(load.position, false)), new Node { power = load.power });
                    }
                }
                treeline.nodes.Add(0, new Node());
            }

            foreach (var treeline in treelines)
            {
                treeline.nodes.ElementAt(0).Value.distance = treeline.nodes.ElementAt(0).Key;
                for (int i = 1; i < treeline.nodes.Count; i++)
                {
                    treeline.nodes.ElementAt(i).Value.distance = treeline.nodes.ElementAt(i).Key - treeline.nodes.ElementAt(i - 1).Key;
                }
            }

            var root = (from l in treelines where l.ischild == false select l).First();

            return EquelNode(root).moment;
        }

        static Node EquelNode(BranchNode root)
        {
            double currmoment = 0;
            for (int i = root.nodes.Count - 1; i >= 1; i--)
            {
                if (root.nodes.ElementAt(i - 1).Value is BranchNode)
                {
                    var equelnode = EquelNode(root.nodes.ElementAt(i - 1).Value as BranchNode);
                    root.nodes.ElementAt(i - 1).Value.power = equelnode.power;
                    root.nodes.ElementAt(i - 1).Value.moment = equelnode.moment;
                }
                root.nodes.ElementAt(i - 1).Value.power += root.nodes.ElementAt(i).Value.power;
                currmoment += root.nodes.ElementAt(i).Value.power * root.nodes.ElementAt(i).Value.distance / 1000000;
                if (currmoment > root.nodes.ElementAt(i - 1).Value.moment)
                    root.nodes.ElementAt(i - 1).Value.moment = currmoment;
                else
                    currmoment = root.nodes.ElementAt(i - 1).Value.moment;
            }
            return root.nodes.ElementAt(0).Value;
        }

        public static void ReversePoly(Polyline poly)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                Polyline p = (Polyline)tr.GetObject(poly.ObjectId, OpenMode.ForWrite);
                p.ReverseCurve();
                tr.Commit();
            }
        }

        public static double Distance(Point3d p1, Point3d p2)
        {
            return Math.Sqrt(Math.Pow((p1.X - p2.X), 2) + Math.Pow((p1.Y - p2.Y), 2));
        }
        
        public void Initialize()
        {          
        }

        public void Terminate()
        {           
        }
    }
}
