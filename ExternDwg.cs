using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

namespace C3D_table_export
{
    public class ExternDwg
    {
        public static void CopySpaceToExtDWG()
        {
            // get the working database (in AutoCAD)
            Database sourceDb = Application.DocumentManager.MdiActiveDocument.Database;

            try
            {
                // create a new destination database
                using (Database destDb = new Database(true, true))
                {
                    // get the model space object ids for both dbs
                    ObjectId sourceMsId = SymbolUtilityServices.GetBlockModelSpaceId(sourceDb);
                    ObjectId destDbMsId = SymbolUtilityServices.GetBlockModelSpaceId(destDb);

                    // now create an array of object ids to hold the source objects to copy
                    ObjectIdCollection sourceIds = new ObjectIdCollection();

                    // open the sourceDb ModelSpace (current autocad dwg)
                    using (BlockTableRecord ms =  sourceMsId.Open(OpenMode.ForRead) as BlockTableRecord)

                        // loop all the entities and record their ids
                        foreach (ObjectId id in ms)
                            sourceIds.Add(id);

                    // next prepare to deepclone the recorded ids to the destdb
                    IdMapping mapping = new IdMapping();

                    // now clone the objects into the destdb
                    sourceDb.WblockCloneObjects(sourceIds, destDbMsId, mapping, DuplicateRecordCloning.Replace, false);

                    destDb.SaveAs("c:\\temp\\CopyTest.dwg", DwgVersion.Current);
                }

            }

            catch (System.Exception eXP)

            {

                System.Windows.Forms.MessageBox.Show(eXP.ToString());

            }
        }
    }
}
