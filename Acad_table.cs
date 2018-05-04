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
#endregion


namespace C3D_table_export
{
    public class Acad_table
    {
        public void Select()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            List<TableData> list = new List<TableData>();
            List<string> values;

            PromptEntityOptions peo = new PromptEntityOptions("\nSelect table");
            peo.SetRejectMessage("\nMust be a table.");
            peo.AddAllowedClass(typeof(Table), false);
            peo.AllowNone = false;
            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK)
                return;

            ObjectId tabId = per.ObjectId;

            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var tab = tr.GetObject(tabId, OpenMode.ForRead) as Table;
                using (var mt = new MText())
                {
                    for (int r = 0; r < tab.Rows.Count; r++)
                    {
                        values = new List<string>();
                        for (int c = 0; c < tab.Columns.Count; c++)
                        {
                            var cell = tab.Cells[r, c];
                            values.Add(cell.TextString);
                            //mt.Contents = cell.TextString;
                        }

                        list.Add(new TableData
                        {
                            Line = r.ToString(), Values = values
                        });
                    }
                }
            }

            ExportExcel.ExportToExcel(list);
        }


    }

    public class TableData
    {
        public string Line { get; set; }
        public List<string> Values { get; set; }
    }
}