// Copyright (c) 2013, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using TestDriverCodePack;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;

namespace TestDriver
{
    //-------------------------------------------------------------------------------------------------------------------------
    //
    // Prerequisites:
    //      PropertyHandler registered and set up on .txt extension
    //      Our bitness (32/64) matches bitness of FileMeta setup.  32-bit on 64-bit Windows should work.
    //
    //-------------------------------------------------------------------------------------------------------------------------

    // Write a single property using API COde Pack (which uses the Property Handler), and read it back directly using the Property Handler
    public class RoundTrip1 : Test
    {
        public override string Name { get { return "Write API Code Pack, read Property handler"; } }

        public override bool RunBody(State state)
        {
            RequirePropertyHandlerRegistered();
            RequireTxtProperties();

            const string cval = "acomment!!";
            string propertyName = "System.Comment";

            //Create a temp file to put metadata on
            string fileName = CreateFreshFile(1);

            // Use API Code Pack to set the value
            IShellProperty prop = ShellObject.FromParsingName(fileName).Properties.GetProperty(propertyName);
            (prop as ShellProperty<string>).Value = cval;
            string svalue = null;

            var handler = new CPropertyHandler();
            handler.Initialize(fileName, 0);

            PropVariant value = new PropVariant();
            handler.GetValue(new TestDriverCodePack.PropertyKey(new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9"), 6), value);
            svalue = (string)value.Value;

            Marshal.ReleaseComObject(handler);  // preempt GC for CCW

            File.Delete(fileName);  // only works if all have let go of the file

            return svalue == cval;
        }
    }

    // Write a single property directly using the Property Handler, and read it back using API COde Pack (which uses the Property Handler)
    public class RoundTrip2 : Test
    {
        public override string Name { get { return "Write Property handler, read API Code Pack"; } }

        public override bool RunBody(State state)
        {
            RequirePropertyHandlerRegistered();
            RequireTxtProperties();

            const string cval = "bcomment??";
            string propertyName = "System.Comment";

            //Create a temp file to put metadata on
            string fileName = CreateFreshFile(1);

            var handler = new CPropertyHandler();
            handler.Initialize(fileName, 0);

            PropVariant value = new PropVariant(cval);
            handler.SetValue(new TestDriverCodePack.PropertyKey(new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9"), 6), value);
            handler.Commit();

            Marshal.ReleaseComObject(handler);  // preempt GC for CCW

            // Use API Code Pack to read the value
            IShellProperty prop = ShellObject.FromParsingName(fileName).Properties.GetProperty(propertyName);
            object oval = prop.ValueAsObject;

            File.Delete(fileName);  // only works if all have let go of the file

            return (string)oval == cval;
        }
    }

    // Write a single property using DSOFile (which does not use the Property Handler), and read it back directly using the Property Handler
    // This only does anything in 32-bit, since DSOFile only ships as 32-bit
    public class RoundTrip3 : Test
    {
        public override string Name { get { return "Write DSOFile, read Property handler"; } }

        public override bool RunBody(State state)
        {
            RequirePropertyHandlerRegistered();
            RequireTxtProperties();

#if x86

            const string cval = "ccomment***";

            //Create a temp file to put metadata on
            string fileName = CreateFreshFile(1);

            // Use DSOFile to set the value
            var dso = new DSOFile.OleDocumentProperties();
            dso.Open(fileName);
            var sum = dso.SummaryProperties;
            sum.Comments = cval;
            dso.Save();
            dso.Close();
            Marshal.ReleaseComObject(dso);  // preempt GC for CCW

            string svalue = null;

            var handler = new CPropertyHandler();
            handler.Initialize(fileName, 0);

            PropVariant value = new PropVariant();
            handler.GetValue(new TestDriverCodePack.PropertyKey(new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9"), 6), value);
            svalue = (string)value.Value;

            Marshal.ReleaseComObject(handler);  // preempt GC for CCW

            File.Delete(fileName);  // only works if all have let go of the file

            return svalue == cval;
#else
            state.RecordEntry("Test skipped because DSOFile is 32-bit only");
            return true;
#endif
        }
    }

    // Write and read a non-Latin property value using the Property Handler
    public class RoundTrip4 : Test
    {
        public override string Name { get { return "Write & read non-Latin values using Property handler"; } }

        public override bool RunBody(State state)
        {
            RequirePropertyHandlerRegistered();
            RequireTxtProperties();

            const string cval1 = "할말있어";
            string[] cval2 = { "hello", "Приветствия"};

            //Create a temp file to put metadata on
            string fileName = CreateFreshFile(1);

            var handler = new CPropertyHandler();
            handler.Initialize(fileName, 0);

            PropVariant value1 = new PropVariant(cval1);
            PropVariant value2 = new PropVariant(cval2);
            // System.Comment
            handler.SetValue(new TestDriverCodePack.PropertyKey(new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9"), 6), value1);
            // System.Category
            handler.SetValue(new TestDriverCodePack.PropertyKey(new Guid("D5CDD502-2E9C-101B-9397-08002B2CF9AE"), 2), value2);
            handler.Commit();

            PropVariant getvalue1 = new PropVariant();
            PropVariant getvalue2 = new PropVariant();
            handler.GetValue(new TestDriverCodePack.PropertyKey(new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9"), 6), getvalue1);
            handler.GetValue(new TestDriverCodePack.PropertyKey(new Guid("D5CDD502-2E9C-101B-9397-08002B2CF9AE"), 2), getvalue2);
            string result1 = (string)getvalue1.Value;
            string[] result2 = (string[])getvalue2.Value;

            Marshal.ReleaseComObject(handler);  // preempt GC for CCW

            File.Delete(fileName);  // only works if all have let go of the file

            return result1 == cval1 && result2[0] == cval2[0] && result2[1] == cval2[1];
        }
    }

    // Test read/write patterns using the Property Handler
    public class RoundTrip5 : Test
    {
        public override string Name { get { return "Read/write patterns using the Property Handler"; } }

        public override bool RunBody(State state)
        {
            RequirePropertyHandlerRegistered();
            RequireTxtProperties();

            string propertyName1 = "System.Comment";
            const string cval1 = "comment";
            const string cval1a = "commenta";
            Guid format1 = new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9");
            int id1 = 6;

            string propertyName2 = "System.Title";
            const string cval2 = "title";
            const string cval2a = "titlea";
            Guid format2 = new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9");
            int id2 = 2;

            //Create a temp file to put metadata on
            string fileName = CreateFreshFile(1);

            // Use API Code Pack to set the values
            IShellProperty prop1 = ShellObject.FromParsingName(fileName).Properties.GetProperty(propertyName1);
            (prop1 as ShellProperty<string>).Value = cval1;
            IShellProperty prop2 = ShellObject.FromParsingName(fileName).Properties.GetProperty(propertyName2);
            (prop2 as ShellProperty<string>).Value = cval2;

            var handler = new CPropertyHandler();
            handler.Initialize(fileName, 0);

            // Read the values with the Property Handler
            PropVariant getvalue1 = new PropVariant();
            PropVariant getvalue2 = new PropVariant();
            handler.GetValue(new TestDriverCodePack.PropertyKey(format1, id1), getvalue1);
            handler.GetValue(new TestDriverCodePack.PropertyKey(format2, id2), getvalue2);
            string result1 = (string)getvalue1.Value;
            string result2 = (string)getvalue2.Value;

            if (result1 != cval1 || result2 != cval2)
                return false;

            // Set the values with the Property Handler
            PropVariant value1 = new PropVariant(cval1a);
            PropVariant value2 = new PropVariant(cval2a);
            // System.Comment
            handler.SetValue(new TestDriverCodePack.PropertyKey(format1, id1), value1);
            // System.Category
            handler.SetValue(new TestDriverCodePack.PropertyKey(format2, id2), value2);
            handler.Commit();

            // Read the updated values with the Property Handler
            getvalue1 = new PropVariant();
            getvalue2 = new PropVariant();
            handler.GetValue(new TestDriverCodePack.PropertyKey(format1, id1), getvalue1);
            handler.GetValue(new TestDriverCodePack.PropertyKey(format2, id2), getvalue2);
            result1 = (string)getvalue1.Value;
            result2 = (string)getvalue2.Value;

            Marshal.ReleaseComObject(handler);  // preempt GC for CCW

            File.Delete(fileName);  // only works if all have let go of the file

            return (result1 == cval1a && result2 == cval2a);
        }
    }
}
