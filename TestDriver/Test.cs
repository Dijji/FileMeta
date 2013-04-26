// Copyright (c) 2013, Dijii, and released under the Common Public License.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;

namespace TestDriver
{
    public abstract class Test
    {
        const string OurPropertyHandlerGuid64 = "{D06391EE-2FEB-419B-9667-AD160D0849F3}";
        const string OurPropertyHandlerGuid32 = "{60211757-EF87-465e-B6C1-B37CF98295F9}";
        const string OurContextHandlerGuid64 = "{28D14D00-2D80-4956-9657-9D50C8BB47A5}";
        const string OurContextHandlerGuid32 = "{DA38301B-BE91-4397-B2C8-E27A0BD80CC5}";

        private static Nullable<bool> isOurPropertyHandlerRegistered = null;
        private static Nullable<bool> isOurContextHandlerRegistered = null;
        private static Nullable<bool> isTxtPropertyHandlerRegistered = null;

        public abstract string Name { get; }

        public bool Run(State state)
        {
            bool good = false;
            try
            {
                good = RunBody(state);
            }
            catch (System.Exception e)
            {
                state.RecordEntry(e.ToString());
            }
            state.RecordResult(Name, good);

            return good;
        }

        public abstract bool RunBody(State state);

#if x64
        private static string OurPropertyHandlerGuid { get { return OurPropertyHandlerGuid64; } }
        private static string OurContextHandlerGuid { get { return OurContextHandlerGuid64; } }
#elif x86
        private static string OurPropertyHandlerGuid { get { return OurPropertyHandlerGuid32; } }
        private static string OurContextHandlerGuid { get { return OurContextHandlerGuid32; } }
#endif

        public static bool IsOurPropertyHandlerRegistered
        {
            get
            {
                // Cache the answer to avoid pounding on the registry
                if (isOurPropertyHandlerRegistered == null)
                    isOurPropertyHandlerRegistered = (Registry.ClassesRoot.OpenSubKey(@"CLSID\" + OurPropertyHandlerGuid, false) != null);

                return (bool)isOurPropertyHandlerRegistered;
            }
        }

        public static bool IsOurContexrHandlerRegistered
        {
            get
            {
                // Cache the answer to avoid pounding on the registry
                if (isOurContextHandlerRegistered == null)
                    isOurContextHandlerRegistered = (Registry.ClassesRoot.OpenSubKey(@"CLSID\" + OurContextHandlerGuid, false) != null);

                return (bool)isOurContextHandlerRegistered;
            }
        }

        protected void RequirePropertyHandlerRegistered()
        {
            if (!IsOurPropertyHandlerRegistered)
                throw new System.Exception("Prequisite: Property Handler must be registered");
        }

        protected void RequireContextHandlerRegistered()
        {
            if (!IsOurContexrHandlerRegistered)
                throw new System.Exception("Prequisite: Context Menu Handler must be registered");
        }

        protected void RequireTxtProperties()
        {
            if (isTxtPropertyHandlerRegistered == null)
            {
                isTxtPropertyHandlerRegistered = false;
                var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers\.txt", false);

                if (key != null)
                {
                    string sval = key.GetValue(null) as string;
                    if (sval != null && sval == OurPropertyHandlerGuid)
                        isTxtPropertyHandlerRegistered = true;
                }
            }

            if (!(bool)isTxtPropertyHandlerRegistered)
                throw new System.Exception("Prequisite: .txt extension must be set to use our Property Handler");
        }

        protected string CreateFreshFile(int index)
        {
            //Create a temp file to put metadata on
            string fileName = Path.GetTempPath() + "test" + index.ToString() + ".txt";

            // Need delete as ovverwrite won't clear alternate streams
            if (File.Exists(fileName))
                File.Delete(fileName);
            File.Create(fileName).Close(); // if no close, property set fails with lock error

            return fileName;
        }

        protected void RenameWithDelete(string from, string to)
        {
            if (File.Exists(to))
                File.Delete(to);
            File.Move(from, to);
        }

        protected string MetadataFileName(string fileName)
        {
            return fileName + ".metadata.xml"; 
        }
    }
}
