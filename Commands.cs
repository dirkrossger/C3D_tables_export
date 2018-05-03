﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

[assembly: CommandClass(typeof(C3D_table_export.Commands))]


namespace C3D_table_export
{
    class Commands
    {
        [CommandMethod("command2")]
        public void Command1()
        {
            C3D_table tabl = new C3D_table();
            tabl.Select("\nSelect Material tables to extract: ");
        }
    }

    public class C3D_sect_mtrl : IExtensionApplication
    {
        [CommandMethod("info")]
        public void Initialize()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;

            ed.WriteMessage("\n-> Get Area OF Material Solids: command2");

        }

        public void Terminate()
        {
        }
    }

}