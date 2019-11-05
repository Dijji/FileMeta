// Copyright (c) 2016, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace TestDriverAssoc
{
    class TestCommandLine
    {
        public enum WindowsErrorCode
        {
            ERROR_FILE_NOT_FOUND = 2,
            ERROR_ACCESS_DENIED = 5,
            ERROR_INVALID_PARAMETER = 87,
            ERROR_XML_PARSE_ERROR = 1465,
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="obj"></param>
        public static void Run(Object obj)
        {
            State state = (State)obj;
            state.TestCounter = 1;
            bool overallPass = true;

            // Backup any saved state file because some add operations update it
            if (!Common.SetupProfileForTestRun(state))
                return;

            state.AddReport("");
            state.AddReport(String.Format("Run {0}", state.RunCounter));
            state.AddReport("");

            try
            {
                bool pass = true;

                // This is an internal test to verify that we are writing and reading all the registry settings that we use accurately
                // Otherwise, we might get false negatives when we run our real tests
                state.AddReport(String.Format("#{0}: Registry settings round trips", state.TestCounter));

                RoundTrip(state, "V13BuiltIn", ref Const.V13BuiltIn, ref pass);
                RoundTrip(state, "V13Custom", ref Const.V13CustomTest, ref pass);
                RoundTrip(state, "V14BuiltIn", ref Const.V14BuiltIn, ref pass);
                RoundTrip(state, "V14Custom", ref Const.V14CustomTest, ref pass);
                RoundTrip(state, "V14Extended", ref Const.V14ExtendedBmp, ref pass);
                RoundTrip(state, "V14UnExtended", ref Const.V14UnExtended, ref pass);
                RoundTrip(state, "V15BuiltIn", ref Const.V15BuiltIn, ref pass);
                RoundTrip(state, "V15Custom", ref Const.V15CustomTest, ref pass);
                RoundTrip(state, "V15CustomOther32", ref Const.V15CustomTestOther32, ref pass);
                RoundTrip(state, "V15CustomOther64", ref Const.V15CustomTestOther64, ref pass);
                RoundTrip(state, "V15Extended", ref Const.V15ExtendedBmp, ref pass);
#if x64
                RoundTrip(state, "V15InitialOther32", ref Const.V15InitialOther32, ref pass);
#endif
                RoundTrip(state, "V15UnExtended", ref Const.V15UnExtended, ref pass);

                state.AddReport(String.Format("#{0}: {1}", state.TestCounter++, pass ? "Passed" : "Failed"));
                overallPass &= pass;

                // Test command line utility's ability to remove handlers set up by various versions
                // Since upgrading is always a matter of remove then re-add, this is all we need to do to check old version handling
                pass = true;
                state.AddReport("");
                state.AddReport(String.Format("#{0}: Remove handlers set up by various versions", state.TestCounter));

                Remove(state, "Version 1.3 with Simple profile", "V13BuiltIn", ref Const.V13BuiltIn, ref pass);
                Remove(state, "Version 1.3 with custom profile", "V13Custom", ref Const.V13CustomTest, ref pass);
                Remove(state, "Version 1.4 with Simple profile", "V14BuiltIn", ref Const.V14BuiltIn, ref pass);
                Remove(state, "Version 1.4 with custom profile", "V14Custom", ref Const.V14CustomTest, ref pass);
                Remove(state, "Version 1.4 with extended handler", "V14Extended", ref Const.V14ExtendedBmp, ref pass, Const.V14UnExtended);
                Remove(state, "Version 1.5 with Simple profile", "V15BuiltIn", ref Const.V15BuiltIn, ref pass);
                Remove(state, "Version 1.5 with custom profile", "V15Custom", ref Const.V15CustomTest, ref pass);
#if x64
                Remove(state, "Version 1.5 with custom profile and existing 32 bit handler", "V15CustomOther32", ref Const.V15CustomTestOther32, ref pass, Const.V15InitialOther32);
                Remove(state, "Version 1.5 with custom profile and existing 64 bit handler", "V15CustomOther64", ref Const.V15CustomTestOther64, ref pass, Const.V15InitialOther64);
#endif
                Remove(state, "Version 1.5 with extended handler", "V15Extended", ref Const.V15ExtendedBmp, ref pass, Const.V15UnExtended);
                Remove(state, "Version 1.5 with extended handler and CLSID settings", "V15ExtendedClsid", ref Const.V15ExtendedBmpClsid, ref pass, Const.V15UnExtendedClsid);
                Remove(state, "Version 1.5 with extended handler and both settings", "V15ExtendedBoth", ref Const.V15ExtendedBmpBoth, ref pass, Const.V15UnExtendedBoth);

                state.AddReport(String.Format("#{0}: {1}", state.TestCounter++, pass ? "Passed" : "Failed"));
                overallPass &= pass;

                // Test command line utility's ability to add handlers with various arguments
                pass = true;
                state.AddReport("");
                state.AddReport(String.Format("#{0}: Add handlers with various arguments", state.TestCounter));

                Add(state, "Extension does not exist, Simple profile", "V15BuiltIn", "-p=simple", ref Const.V15BuiltIn, ref pass);
                RegState.CreateMinimalExtension(Const.TestExt);
                Add(state, "Minimal extension, Simple profile", "V15BuiltIn", "-p=simple", ref Const.V15BuiltIn, ref pass);
                Add(state, "Extension does not exist, custom profile 'test'", "V15CustomTest", "-p=test -d=SavedState.xml", ref Const.V15CustomTest, ref pass);
#if x64
                Add(state, "Extension does not exist, 32 bit handler does, custom profile 'test'", "V15CustomTestOther32", "-p=test -d=SavedState.xml", ref Const.V15CustomTestOther32, ref pass, Const.V15InitialOther32);
                Add(state, "Extension does not exist, 64 bit handler does, custom profile 'test'", "V15CustomTestOther64", "-p=test -d=SavedState.xml", ref Const.V15CustomTestOther64, ref pass, Const.V15InitialOther64);
#endif
                Common.ClearProfile();
                Add(state, "Extend existing .bmp property handler", "V15ExtendedBmp", "-p=.bmp -d=SavedState.xml", ref Const.V15ExtendedBmp, ref pass, Const.V15UnExtended);
                if (!TestGUI.VerifySavedState("SavedStateBmpOnly.xml"))
                {
                    state.AddReport(String.Format("Add did not produce the expected saved state file for {0}", Const.TestExt));
                    pass = false;
                }
                Add(state, "Extend existing .bmp property handler with CLSID settings", "V15ExtendedBmpClsid", "-p=.bmp -d=SavedState.xml", ref Const.V15ExtendedBmpClsid, ref pass, Const.V15UnExtendedClsid);
                Add(state, "Extend existing .bmp property handler with both settings", "V15ExtendedBmpBoth", "-p=.bmp -d=SavedState.xml", ref Const.V15ExtendedBmpBoth, ref pass, Const.V15UnExtendedBoth);
                Common.ResetProfile();

                state.AddReport(String.Format("#{0}: {1}", state.TestCounter++, pass ? "Passed" : "Failed"));
                overallPass &= pass;

                // Test command line utility's ability to handle various error conditions
                pass = true;
                state.AddReport("");
                state.AddReport(String.Format("#{0}: Test various error conditions", state.TestCounter));

                Error(state, "No arguments at all", "", WindowsErrorCode.ERROR_INVALID_PARAMETER, ref pass, true);
                Error(state, "Bad command", "-x", WindowsErrorCode.ERROR_INVALID_PARAMETER, ref pass, true);
                Error(state, "Two commands", "-a -r", WindowsErrorCode.ERROR_INVALID_PARAMETER, ref pass, true);
                Error(state, "Remove without an extension", "-r", WindowsErrorCode.ERROR_INVALID_PARAMETER, ref pass, true);
                Error(state, "Remove with a bad extension", "-r nosuch", WindowsErrorCode.ERROR_INVALID_PARAMETER, ref pass, true);
                Error(state, "Remove with a non-existent extension", "-r", WindowsErrorCode.ERROR_INVALID_PARAMETER, ref pass, false);
                Error(state, "Add without an extension", "-a", WindowsErrorCode.ERROR_INVALID_PARAMETER, ref pass, true);
                Error(state, "Add without a profile", "-a", WindowsErrorCode.ERROR_INVALID_PARAMETER, ref pass, false);
                Error(state, "Add with a bad profile", "-a -p=nosuch", WindowsErrorCode.ERROR_INVALID_PARAMETER, ref pass, false);
                Error(state, "Add with a data file and a bad profile", "-a -p=nosuch -d=SavedState.xml", WindowsErrorCode.ERROR_INVALID_PARAMETER, ref pass, false);

                state.AddReport(String.Format("#{0}: {1}", state.TestCounter++, pass ? "Passed" : "Failed"));
                overallPass &= pass;
            }
            catch (Exception ex)
            {
                state.AddReport(String.Format("Run failed with unexpected exception '{0}'", ex.Message));
            }
            finally
            {
                Common.RestoreProfileAfterTestRun(state);
            }

            state.AddReport("");
            state.AddReport(String.Format("Run {0} completed {1}", state.RunCounter++, overallPass ? "with no failures" : "with some failures"));
        }

        private static void RoundTrip(State state, string name, ref RegState source, ref bool pass)
        {
            source.Zap(Const.TestExt);
            var read = new RegState();
            read.Read(Const.TestExt);
            if (read != source)
            {
                state.AddReport(String.Format("Round-trip failed for {0}", name));
                pass = false;
            }

            RegState.Wipe(Const.TestExt);

            var clear = new RegState();
            clear.Read(Const.TestExt);
            if (clear != new RegState())
            {
                state.AddReport(String.Format("Wipe failed for {0}", name));
                pass = false;
            }
        }

        public static void Remove(State state, string description, string name, ref RegState source, ref bool pass, RegState? final = null)
        {
            state.AddReport(description);
            source.Zap(Const.TestExt);
            var result = InvokeFileMetaAssoc(state, "-r");
            if (result != 0)
            {
                state.AddReport(String.Format("FileMetaAssoc -r failed for {0}", name));
                pass = false;
            }

            var outcome = new RegState();
            outcome.Read(Const.TestExt);

            if (final == null)
            {
                // Usually, the expected final state is an empty registry
                if (outcome != new RegState())
                {
                    state.AddReport(String.Format("Remove did not completely clean up for {0}", name));
                    pass = false;
                }
            }
            else
            {
                // If the final state was specified, verify that we got it
                if (outcome != final)
                {
                    state.AddReport(String.Format("Remove did not produce the expected final registry state for {0}", name));
                    pass = false;
                }
            }

            // Clean up after ourselves
            RegState.Wipe(Const.TestExt);
        }

        public static void Add(State state, string description, string name, string args, ref RegState final, ref bool pass, RegState? initial = null, bool noWipe = false)
        {
            state.AddReport(description);
            if (initial != null)
                ((RegState)initial).Zap(Const.TestExt);

            var result = InvokeFileMetaAssoc(state, "-a " + args);
            if (result != 0)
            {
                state.AddReport(String.Format("FileMetaAssoc -a failed for {0}", name));
                pass = false;
            }

            if (pass)
            {
                var outcome = new RegState();
                outcome.Read(Const.TestExt);

                // Verify that we got the final state
                if (outcome != final)
                {
                    state.AddReport(String.Format("Add did not produce the expected final registry state for {0}", name));
                    pass = false;
                }
            }

            // Clean up after ourselves
            if (!noWipe)
                RegState.Wipe(Const.TestExt);
        }

        private static void Error(State state, string description, string args, WindowsErrorCode error, ref bool pass, bool noExtension)
        {
            state.AddReport(description);
            var result = InvokeFileMetaAssoc(state, args, true, noExtension);

            if (result != (int)error)
            {
                state.AddReport(String.Format("FileMetaAssoc returned {0} instead of the expected error code {1}", result, error));
                pass = false;
            }
        }

        private static int InvokeFileMetaAssoc(State state, string args, bool noErrorReport = false, bool noExtension = false)
        {
            Process p = new Process();
            string line;

            // Redirect the output stream of the child process.
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.CreateNoWindow = true;
            p.StartInfo.FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"File Metadata\FileMetaAssoc.exe");
            if (noExtension)
                p.StartInfo.Arguments = args;
            else
                p.StartInfo.Arguments = String.Format("{0} \"{1}\"", args, Const.TestExt);
            p.Start();
            while (true)
            {
                if (p.StandardOutput.EndOfStream && p.HasExited)
                    break;
                else if (!p.StandardOutput.EndOfStream)
                {
                    line = p.StandardOutput.ReadLine();
                    state.AddReport(line);
                }
                System.Threading.Thread.Sleep(100);  // Be polite
            }
            string error = p.StandardError.ReadToEnd().Replace("\r", "").Replace("\n", "");
            if (error != null && error.Length > 0)
                state.AddReport(error);
            if (p.ExitCode != 0 && !noErrorReport)
                state.AddReport(String.Format("FileMetaAssoc returned error code {0}", p.ExitCode));
            return p.ExitCode;
        }
    }
}
