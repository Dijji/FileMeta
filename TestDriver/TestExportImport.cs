// Copyright (c) 2013, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;
using System.Runtime.InteropServices;

namespace TestDriver
{
    public class ExportImport1 : Test
    {
        public override string Name { get { return "Export 1 property, import to different file"; } }

        public override bool RunBody(State state)
        {
            RequirePropertyHandlerRegistered();
            RequireContextHandlerRegistered();
            RequireTxtProperties();

            const string cval = "exp-imp";
            string propertyName = "System.Comment";

            //Create a temp file to put metadata on
            string fileName = CreateFreshFile(1);

            // Use API Code Pack to set the value
            IShellProperty prop = ShellObject.FromParsingName(fileName).Properties.GetProperty(propertyName);
            (prop as ShellProperty<string>).Value = cval;

            // Set up ContextHandler and tell it about the target file
            var handler = new CContextMenuHandler();
            var dobj = new DataObject();
            dobj.SetFileDropList(new StringCollection{fileName});
            handler.Initialize(new IntPtr(0), dobj, 0);
            handler.QueryContextMenu(0, 0, 0, 0, 0); // This fails, but  that's ok

            //export the metadata
            CMINVOKECOMMANDINFOEX cmd = new CMINVOKECOMMANDINFOEX();
            cmd.lpVerb = new IntPtr((int)ContextMenuVerbs.Export);
            var cw = new CommandWrapper(cmd);
            handler.InvokeCommand(cw.Ptr);

            // Create a 2nd temp file for import
            string fileName2 = CreateFreshFile(2);

            Marshal.ReleaseComObject(handler);  // preempt GC for CCW

            // rename metadata file
            RenameWithDelete(MetadataFileName(fileName), MetadataFileName(fileName2));

            // Get a new handler and import
            handler = new CContextMenuHandler();
            dobj = new DataObject();
            dobj.SetFileDropList(new StringCollection { fileName2 });
            handler.Initialize(new IntPtr(0), dobj, 0);
            handler.QueryContextMenu(0, 0, 0, 0, 0); // This fails, but  that's ok

            cmd = new CMINVOKECOMMANDINFOEX();
            cmd.lpVerb = new IntPtr((int)ContextMenuVerbs.Import);
            cw = new CommandWrapper(cmd);
            handler.InvokeCommand(cw.Ptr);

            Marshal.ReleaseComObject(handler);  // preempt GC for CCW

            // Use API Code Pack to read the value
            prop = ShellObject.FromParsingName(fileName2).Properties.GetProperty(propertyName);
            object oval = prop.ValueAsObject;

            // Clean up files - checks if they have been released, too
            File.Delete(fileName);
            File.Delete(fileName2);
            File.Delete(MetadataFileName(fileName2));  

            // good if original value has made it round
            return cval == (string)oval;
        }
    }

    public class ExportImport2 : Test
    {
        public override string Name { get { return "Export non-Latin properties, import to different file"; } }

        public override bool RunBody(State state)
        {
            RequirePropertyHandlerRegistered();
            RequireContextHandlerRegistered();
            RequireTxtProperties();

            string cval1 = "할말있어";
            string[] cval2 = { "hello", "Приветствия" };
            string cval3 = "title";
            string propertyName1 = "System.Comment";
            string propertyName2 = "System.Category";
            string propertyName3 = "System.Title";

            //Create a temp file to put metadata on
            string fileName = CreateFreshFile(1);

            // Use API Code Pack to set the values
            IShellProperty prop1 = ShellObject.FromParsingName(fileName).Properties.GetProperty(propertyName1);
            (prop1 as ShellProperty<string>).Value = cval1;
            IShellProperty prop2 = ShellObject.FromParsingName(fileName).Properties.GetProperty(propertyName2);
            (prop2 as ShellProperty<string[]>).Value = cval2;
            IShellProperty prop3 = ShellObject.FromParsingName(fileName).Properties.GetProperty(propertyName3);
            (prop3 as ShellProperty<string>).Value = cval3;

            // Set up ContextHandler and tell it about the target file
            var handler = new CContextMenuHandler();
            var dobj = new DataObject();
            dobj.SetFileDropList(new StringCollection { fileName });
            handler.Initialize(new IntPtr(0), dobj, 0);
            handler.QueryContextMenu(0, 0, 0, 0, 0); // This fails, but  that's ok

            //export the metadata
            CMINVOKECOMMANDINFOEX cmd = new CMINVOKECOMMANDINFOEX();
            cmd.lpVerb = new IntPtr((int)ContextMenuVerbs.Export);
            var cw = new CommandWrapper(cmd);
            handler.InvokeCommand(cw.Ptr);

            // Create a 2nd temp file for import
            string fileName2 = CreateFreshFile(2);

            Marshal.ReleaseComObject(handler);  // preempt GC for CCW

            // rename metadata file
            RenameWithDelete(MetadataFileName(fileName), MetadataFileName(fileName2));

            // Get a new handler and import
            handler = new CContextMenuHandler();
            dobj = new DataObject();
            dobj.SetFileDropList(new StringCollection { fileName2 });
            handler.Initialize(new IntPtr(0), dobj, 0);
            handler.QueryContextMenu(0, 0, 0, 0, 0); // This fails, but  that's ok

            cmd = new CMINVOKECOMMANDINFOEX();
            cmd.lpVerb = new IntPtr((int)ContextMenuVerbs.Import);
            cw = new CommandWrapper(cmd);
            handler.InvokeCommand(cw.Ptr);

            Marshal.ReleaseComObject(handler);  // preempt GC for CCW

            // Use API Code Pack to read the values
            prop1 = ShellObject.FromParsingName(fileName2).Properties.GetProperty(propertyName1);
            string result1 = (string)prop1.ValueAsObject;
            prop2 = ShellObject.FromParsingName(fileName2).Properties.GetProperty(propertyName2);
            string[] result2 = (string[]) prop2.ValueAsObject;
            prop3 = ShellObject.FromParsingName(fileName2).Properties.GetProperty(propertyName3);
            string result3 = (string)prop3.ValueAsObject;

            // Clean up files - checks if they have been released, too
            File.Delete(fileName);
            File.Delete(fileName2);
            File.Delete(MetadataFileName(fileName2));

            // good if original value has made it round
            return cval1 == result1 && result2.Length == 2 && cval2[0] == result2[0] && cval2[1] == result2[1] && cval3 == result3;
        }
    }
}
