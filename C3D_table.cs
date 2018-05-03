using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

#region Autodesk
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
//using Autodesk.AutoCAD.ApplicationServices.Core;
using System.Text.RegularExpressions;
#endregion



namespace C3D_table_export
{
    class C3D_table
    {
        public SelectionSet Select(string message)
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            ed.WriteMessage(message);

            TypedValue[] acTypValAr = new TypedValue[1];
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "AECC_QUANTITY_TAKEOFF_AGGREGATE_TABLE"), 0);

            // Assign the filter criteria to a SelectionFilter object
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);

            // Request for objects to be selected in the drawing area
            PromptSelectionResult acSSPrompt;
            acSSPrompt = ed.GetSelection(acSelFtr);

            // If the prompt status is OK, objects were selected
            if (acSSPrompt.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = acSSPrompt.Value;
                TableValues(acSSet);
                return acSSet;
            }
            else
            {
                Application.ShowAlertDialog("Number of objects selected: 0");
            }

            

            return null;
        }

        public void TableValues(SelectionSet acSSet)
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            //List<Datas> result = new List<Datas>();

            using (Transaction trans = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                try
                {
                    for (int i = 0; i < acSSet.Count; i++)
                    {
                        MaterialSection mtrl = trans.GetObject(acSSet[i].ObjectId, OpenMode.ForRead) as MaterialSection;
                        //result.Add(feat.ObjectId);

                        Autodesk.Civil.DatabaseServices.Table tbl = acSSet[i].ObjectId.GetObject(OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Table;
                       //string stylename = aecObj.StyleName;
                       // string mtrlname = Regex.Split(aecObj.Name, "-")[1];
                       // double area = 0.0;

                       // SectionPointCollection sPts = mtrl.SectionPoints;
                       // Point2dCollection p2 = new Point2dCollection();
                       // foreach (SectionPoint pt in sPts)
                       // {
                       //     p2.Add(new Point2d(pt.Location.X, pt.Location.Y));
                       // }
                        //area = Calculate.Area(p2);

                        //result.Add(new Datas { Material = mtrlname, Area = Math.Round(area, 1) });

                    }
                    //return result;
                }
                catch (System.Exception ex) { }

                trans.Commit();

                //return null;
            }
        }
    }
}
