// Copyright (c) 2013, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Runtime.InteropServices;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;
using TestDriverCodePack;

namespace TestDriver
{
    struct SavedProp
    {
        public string Name;
        public object Value;
    }
    class TestProperties1 : Test
    {
        public override string Name { get { return "Write, read, export and import as many properties as possible"; } }
        private Random random = new Random();
        private const int max16 = 32767;
        private const int max32 = 2147483647;
        private long minTicks = new DateTime(1601, 1, 1).Ticks;  // ensure valid filetime, demanded for some date properties
        private long maxTocks = DateTime.MaxValue.Ticks;
        private List<SavedProp> savedProps;

        public override bool RunBody(State state)
        {
            RequirePropertyHandlerRegistered();
            RequireContextHandlerRegistered();
            RequireTxtProperties();

            state.RecordEntry("Starting mass property setting...");

            //Create a temp file to put metadata on
            string fileName = CreateFreshFile(1);

            savedProps = new List<SavedProp>();

            // Give all writable properties random values, according to their type
            foreach (var propDesc in state.PropertyDescriptions.Where(p => !p.TypeFlags.HasFlag(TestDriverCodePack.PropertyTypeOptions.IsInnate)))
            {
                // These properties don't seem to be writeable.  Don't know why, but they don't appear anyway in the list
                // of usable properties at http://msdn.microsoft.com/en-us/library/windows/desktop/dd561977(v=vs.85).aspx
                if (propDesc.CanonicalName == "System.History.UrlHash" ||
                    propDesc.CanonicalName == "System.DuiControlResource" ||
                    propDesc.CanonicalName == "System.OfflineFiles.CreatedOffline" ||
                    propDesc.CanonicalName == "System.PropList.XPDetailsPanel" ||
                    propDesc.CanonicalName == "System.SDID")
                   continue;

                // Use API Code Pack to set the value, except for strings, because the Code Pack blows when setting strings of length 1 !!
                // Still use Code Pack elsewhere for its nullable type handling
                IShellProperty prop = ShellObject.FromParsingName(fileName).Properties.GetProperty(propDesc.CanonicalName);
                SetPropertyValue(fileName, propDesc, prop);
            }
            state.RecordEntry(String.Format("{0} property values set", savedProps.Count));

            // Go around again, using the Handler directly to read all the values written and then check them
            int errors = GetAndCheckValues(fileName, state);

            state.RecordEntry(String.Format("{0} properties read back, {1} mismatches", savedProps.Count, errors));

            if (errors > 0)
                return false;

            // Use ContextHandler to export all the values
            var contextHandler = new CContextMenuHandler();
            var dobj = new DataObject();
            dobj.SetFileDropList(new StringCollection { fileName });
            contextHandler.Initialize(new IntPtr(0), dobj, 0);

            //export the metadata
            CMINVOKECOMMANDINFOEX cmd = new CMINVOKECOMMANDINFOEX();
            cmd.lpVerb = new IntPtr((int)ContextMenuVerbs.Export);
            var cw = new CommandWrapper(cmd);
            contextHandler.InvokeCommand(cw.Ptr);
            
            Marshal.ReleaseComObject(contextHandler);  // preempt GC for CCW

            // Create new file and import values to it
            string fileName2 = CreateFreshFile(2);

            // rename metadata file
            RenameWithDelete(MetadataFileName(fileName), MetadataFileName(fileName2));
            state.RecordEntry("Metadata exported, starting to import onto new file and check...");

            // Get a new handler and import
            contextHandler = new CContextMenuHandler();
            dobj = new DataObject();
            dobj.SetFileDropList(new StringCollection { fileName2 });
            contextHandler.Initialize(new IntPtr(0), dobj, 0);

            cmd = new CMINVOKECOMMANDINFOEX();
            cmd.lpVerb = new IntPtr((int)ContextMenuVerbs.Import);
            cw = new CommandWrapper(cmd);
            contextHandler.InvokeCommand(cw.Ptr);

            Marshal.ReleaseComObject(contextHandler);  // preempt GC for CCW

            // Go around one last time, reading and checking the imported values
            // We don't use the Code Pack because of it's boolean value bug
            errors = GetAndCheckValues(fileName2, state, false);

            state.RecordEntry(String.Format("{0} properties read back, {1} mismatches", savedProps.Count, errors));

            // Clean up files - checks if they have been released, too
            // Leave files around for analysis if there have been problems
            if (errors == 0)
            {
                File.Delete(fileName);
                File.Delete(fileName2);
                File.Delete(MetadataFileName(fileName2));
            }

            return errors == 0;
        }

        private int GetAndCheckValues(string fileName, State state, bool useCodePack = false)
        {
            int errors = 0;
            CPropertyHandler handler = null;

            if (!useCodePack)
            {
                handler = new CPropertyHandler();
                handler.Initialize(fileName, 0);
            }

            foreach (var saved in savedProps)
            {
                IShellProperty prop = ShellObject.FromParsingName(fileName).Properties.GetProperty(saved.Name);

                object objVal;
                if (useCodePack)
                {
                    objVal = prop.ValueAsObject;
                }
                else
                {
                    var value = new PropVariant();
                    handler.GetValue(new TestDriverCodePack.PropertyKey(prop.PropertyKey.FormatId, prop.PropertyKey.PropertyId), value);
                    objVal = value.Value;
                }
 
                bool bSame = false;
                Type t = objVal != null ? objVal.GetType() : null;
                if (t == typeof(Int16) || t == typeof(Int32) || t == typeof(UInt16) || t == typeof(UInt32) || t == typeof(bool))
                    bSame = objVal.Equals(saved.Value);
                else if (t == typeof(string))
                    bSame = (string)objVal == (string)saved.Value;
                else if (t == typeof(string[]))
                {
                    string[] oss = (string[])objVal;
                    string[] sss = (string[])saved.Value;
                    bSame = true;
                    if (oss.Length == sss.Length)
                    {
                        for (int i = 0; i < oss.Length; i++)
                        {
                            if (oss[i] != sss[i])
                            {
                                bSame = false;
                                break;
                            }
                        }
                    }
                    else
                        bSame = false;
                }
                else if (t == typeof(Int32[]) || t == typeof(UInt32[]))
                {
                    int[] os = (int[])objVal;
                    int[] ss = (int[])saved.Value;
                    bSame = true;
                    if (os.Length == ss.Length)
                    {
                        for (int i = 0; i < os.Length; i++)
                        {
                            if (!os[i].Equals(ss[i]))
                            {
                                bSame = false;
                                break;
                            }
                        }
                    }
                    else
                        bSame = false;
                }
                else if (t == typeof(DateTime))
                {
                    DateTime save = (DateTime)(saved.Value);
                    DateTime read = (DateTime)objVal;

                    // Compare this way because exact ticks don't survive a round trip to text, and legibility
                    // is more useful than serializing just a number of ticks
                    bSame = save.Year == read.Year && save.Month == read.Month && save.Day == read.Day &&
                            save.Hour == read.Hour && save.Minute == read.Minute && save.Second == read.Second &&
                            save.Millisecond == read.Millisecond;
                }
                else
                    bSame = (saved.Value == null && objVal == null);

                if (!bSame)
                {
                    state.RecordEntry(String.Format("Mismatch for property {0}: expected {1}, got {2}", saved.Name,
                                        saved.Value != null ? ToDisplayString(saved.Value) : "Null", objVal != null ? ToDisplayString(objVal) : "Null"));
                    errors++;
                }
            }

            if (!useCodePack)
                Marshal.ReleaseComObject(handler);  // preempt GC for CCW

            return errors;
        }

        private string ToDisplayString(object obj)
        {
            if (obj as Array != null)
            {
                StringBuilder sb = new StringBuilder();
                foreach (var x in obj as Array)
                {
                    sb.Append(x.ToString());
                    sb.Append(";");
                }
                return sb.ToString().TrimEnd(';');
            }
            else
                return obj.ToString();
        } 
        
        private void SetPropertyValue(string fileName, TestDriverCodePack.ShellPropertyDescription propDesc, IShellProperty prop)
        {
            Type t = propDesc.ValueType;
            string s;

            try
            {
                if (t == typeof(string))
                {
                    object obj = GetEnumValueForProperty(propDesc);

                    if (obj != null)
                        s = (string)obj;
                    else
                        s = RandomString();

                    savedProps.Add(new SavedProp { Name = prop.CanonicalName, Value = s });
                    //(prop as ShellProperty<string>).Value = s;

                    // Workaround Code Pack bug with 1 char strings by using PropertyHandler
                    // Have to open and release each time to avoid lock problems - still, it ups the pounding
                    var handler = new CPropertyHandler();
                    handler.Initialize(fileName, 0);

                    PropVariant value = new PropVariant(s);
                    handler.SetValue(new TestDriverCodePack.PropertyKey(prop.PropertyKey.FormatId, prop.PropertyKey.PropertyId), value);
                    handler.Commit();

                    Marshal.ReleaseComObject(handler);  // preempt GC for CCW


                }
                else if (t == typeof(string[]))
                {
                    string[] ss = GetStringArrayValueForProperty(propDesc);

                    savedProps.Add(new SavedProp { Name = prop.CanonicalName, Value = ss });
                    //(prop as ShellProperty<string[]>).Value = ss;

                    // Workaround Code Pack bug with 1 char strings by using PropertyHandler
                    // Have to open and release each time to avoid lock problems - still, it ups the pounding
                    var handler = new CPropertyHandler();
                    handler.Initialize(fileName, 0);

                    PropVariant value = new PropVariant(ss);
                    handler.SetValue(new TestDriverCodePack.PropertyKey(prop.PropertyKey.FormatId, prop.PropertyKey.PropertyId), value);
                    handler.Commit();

                    Marshal.ReleaseComObject(handler);  // preempt GC for CCW
                }
                else if (t == typeof(Int16?) || t == typeof(Int32?) || t == typeof(UInt16?) || t == typeof(UInt32?))
                {
                    object obj = GetEnumValueForProperty(propDesc);

                    if (t == typeof(Int16?))
                    {
                        Int16? val = obj != null ? (Int16?)obj : (Int16?)NullableRandomNumber(-max16, max16);
                        savedProps.Add(new SavedProp { Name = prop.CanonicalName, Value = val });
                        (prop as ShellProperty<Int16?>).Value = val;
                    }
                    else if (t == typeof(Int32?))
                    {
                        Int32? val = obj != null ? (Int32?)obj : (Int32?)NullableRandomNumber(-max32, max32);
                        savedProps.Add(new SavedProp { Name = prop.CanonicalName, Value = val });
                        (prop as ShellProperty<Int32?>).Value = val;
                    }
                    else if (t == typeof(UInt16?))
                    {
                        UInt16? val = obj != null ? (UInt16?)obj : (UInt16?)NullableRandomNumber(max16);
                        savedProps.Add(new SavedProp { Name = prop.CanonicalName, Value = val });
                        (prop as ShellProperty<UInt16?>).Value = val;
                    }
                    else // UInt32?
                    {
                        UInt32? val = obj != null ? (UInt32?)obj : (UInt32?)NullableRandomNumber(max16);
                        savedProps.Add(new SavedProp { Name = prop.CanonicalName, Value = val });
                        (prop as ShellProperty<UInt32?>).Value = val;
                    }
                }
                else if (t == typeof(Int32[]))
                {
                    Int32[] vals = new Int32[4];
                    for (int i = 0; i < 4; i++)
                        vals[i] = RandomNumber(-max32, max32);
                    savedProps.Add(new SavedProp { Name = prop.CanonicalName, Value = vals });
                    (prop as ShellProperty<Int32[]>).Value = vals;
                }
                else if (t == typeof(UInt32[]))
                {
                    UInt32[] vals = new UInt32[4];
                    for (int i = 0; i < 4; i++)
                        vals[i] = (UInt32)RandomNumber(max32);
                    savedProps.Add(new SavedProp { Name = prop.CanonicalName, Value = vals });
                    (prop as ShellProperty<UInt32[]>).Value = vals;
                }
                else if (t == typeof(bool?))
                {
                    int? r = NullableRandomNumber();
                    bool? value = (r == null) ? (bool?)null : (r % 2 == 0);
                    savedProps.Add(new SavedProp { Name = prop.CanonicalName, Value = value });
                    (prop as ShellProperty<bool?>).Value = value;
                }
                else if (t == typeof(DateTime?))
                {
                    DateTime dt = new DateTime((long)(random.NextDouble() * (maxTocks - minTicks) + minTicks));
                    savedProps.Add(new SavedProp { Name = prop.CanonicalName, Value = dt });
                    (prop as ShellProperty<DateTime?>).Value = dt;
                }
                else if (t == typeof(double?))
                {
                    // fails in Code Pack, so skip
                    //(prop as ShellProperty<double>).Value = (double)RandomNumber(max64);
                }
                else if (t == typeof(Int64?))
                {
                    // fails in Code Pack, so skip
                    //(prop as ShellProperty<Int64>).Value = RandomNumber(max64);
                }
                else if (t == typeof(UInt64?))
                {
                    // fails in Code Pack, so skip
                    // (prop as ShellProperty<UInt64>).Value = (UInt64) RandomNumber(max64);
                }
                else if (t == typeof(byte?))
                {
                    // The Code Pack does not support setting these, so skip for now
                    //(prop as ShellProperty<byte>).Value = (byte)RandomNumber(max8);
                }
                else if (t == typeof(byte[]))
                {
                    // The Code Pack does not support setting these, so skip for now
                    // Mostly 128 byte arrays e.g. System.Photo.MakerNote
                    //byte[] bs = new byte[128];
                    //random.NextBytes(bs);
                    //(prop as ShellProperty<byte[]>).Value = bs;
                }
                else if (t == typeof(object) || t == typeof(IntPtr?) || t == typeof(System.Runtime.InteropServices.ComTypes.IStream))
                {
                    // ignore these, they are system artefacts like group header props, and don't appear in settable lists
                }
                else
                    throw new System.Exception("Need " + t.ToString() + " for " + propDesc.CanonicalName);
            }
            catch (System.Exception e)
            {
                throw new System.Exception(String.Format ("Error setting property {0} to '{1}'", 
                    savedProps.Last().Name, ToDisplayString(savedProps.Last().Value)), e);
            }
        }

        private object GetEnumValueForProperty(TestDriverCodePack.ShellPropertyDescription propDesc)
        {
            if (propDesc.PropertyEnumTypes != null && propDesc.PropertyEnumTypes.Count > 0)
            {
                int index = RandomNumber(propDesc.PropertyEnumTypes.Count - 1);

                // Now work out what value to use from the seleccted enum entry
                return GetValueForEnumType(propDesc.PropertyEnumTypes[index]);
            }

            return null;
        }


        private object GetUniqueEnumValueForProperty(TestDriverCodePack.ShellPropertyDescription propDesc, ref bool[] used)
        {
            if (propDesc.PropertyEnumTypes != null && propDesc.PropertyEnumTypes.Count > 0)
            {
                if (used == null)
                    used = new bool[propDesc.PropertyEnumTypes.Count];

                // Ensure that each choice is different - some enums are flags, and insist on unique values
                int usedSoFar = used.Where(f => f == true).Count();
                int index = RandomNumber(propDesc.PropertyEnumTypes.Count - 1 - usedSoFar);
                index += used.Take(index + 1).Where(f => f == true).Count();
                while (used[index])
                    index++;
                used[index] = true;

                // Now work out what string to use from the seleccted enum entry
                return GetValueForEnumType(propDesc.PropertyEnumTypes[index]);
            }

            return null;
        }

        private object GetValueForEnumType(TestDriverCodePack.ShellPropertyEnumType enumType)
        {
            if (enumType.RangeValue != null)
                return enumType.RangeValue;

            else if (enumType.DisplayText != null && enumType.DisplayText.Length > 0)
                // Make the same tweak that ContextManager does, but first, so that is makes our saved value and round-trips work
                return enumType.DisplayText.Replace('\u2013', '-');
            
            else
                throw new System.Exception("puzzling property enum");
        }

        private string[] GetStringArrayValueForProperty(TestDriverCodePack.ShellPropertyDescription propDesc)
        {
            int n = random.Next(2, 5);
            string[] ss = new string[n];
            bool[] used = null;

            for (int i = 0; i < n; i++)
            {
                object obj = GetUniqueEnumValueForProperty(propDesc, ref used);
                if (obj != null)
                    ss[i] = (string) obj;
                else
                    ss[i] = RandomString();
            }

            return ss;
        }

        private int? NullableRandomNumber(int min, int max)
        {
            // 10% of the time, return null
            if (random.Next(10) == 0)
                return null;
            else
                return random.Next(min, max);
        }

        private int? NullableRandomNumber(int max = max16)
        {
            // 10% of the time, return null
            if (random.Next(10) == 0)
                return null;
            else
                return random.Next(max);
        }

        private int RandomNumber(int min, int max)
        {
            return random.Next(min, max);
        }

        private int RandomNumber(int max = max16)
        {
            return random.Next(max);
        }

        private string RandomString(int? size = null)
        {
            StringBuilder builder = new StringBuilder();
            char ch;

            // If no size specified, use 1 to 16 at random
            if (size == null)
                size = random.Next(1, 17); 

            for (int i = 0; i < size; i++)
            {
                // use mixed case
                ch = Convert.ToChar(random.Next(26) + 65 + random.Next(2) * 32);
                builder.Append(ch);
            }
            return builder.ToString();
        }
    }
}
