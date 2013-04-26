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

            //export the metadata
            CMINVOKECOMMANDINFOEX cmd = new CMINVOKECOMMANDINFOEX();
            cmd.lpVerb = new IntPtr((int)ContextMenuVerbs.Export);
            handler.InvokeCommand(cmd);

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

            cmd = new CMINVOKECOMMANDINFOEX();
            cmd.lpVerb = new IntPtr((int)ContextMenuVerbs.Import);
            handler.InvokeCommand(cmd);

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
}
