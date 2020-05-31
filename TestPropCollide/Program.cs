using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using NDesk.Options;
using TestDriverCodePack;

namespace TestPropCollide
{
    class Program
    {
        static readonly string[] helpText = new string[] {
                "",
                "Usage:",
                "",
                "   TestPropCollide.exe [-h] [-d=<target directory>] [-c=<file count>]",
                "                     [-f=<log file>] [-l=<log level>]",
                "",
                "Where:",
                "",
                "   -h, --help",
                "      Display this help",
                "",
                "   -d=<target directory>, --directory=<target directory>",
                "      Directory containing test files, whose properties are to be read and updated",
                "",
                "   -c=<file count>, --filecount=<file count>",
                "      Number of file operations to carry out before test completes. Default is 10000",
                "",
                "   -f=<log file>, --logfile=<log file>",
                "      Full path of the file into which to write the log",
                "",
                "   -l=<log level>, --loglevel=<log level>",
                "      Level of detail for log entries. Between 0 and 3. Default is 1",
                "",
        };

        // Command line arguments and their defaults
        static string targetDirectory = @"d:\TesttagsFlat";
        static long stopCount = 10000;
        static int logLevel = 1;
        static string logFile = @"C:\FileMetaLogs\TestPropCollide.log";

        static int Main(string[] args)
        {
            bool help = false;

            try
            {
                var argParser = new OptionSet() {
                    { "h|?|help", v => {help = v != null; } },
                    { "d|directory=", v => targetDirectory = v },
                    { "c|filecount:", v => stopCount = long.Parse(v) },
                    { "f|logfile=", v => logFile = v },
                    { "l|loglevel:", v => logLevel = int.Parse(v) },
                };
                List<string> extensions = argParser.Parse(args);

                if (help)
                {
                    foreach (var line in helpText)
                        Console.WriteLine(line);
                    return 0;
                }

                File.Delete(logFile);
                Console.WriteLine("Running...");
                Console.WriteLine(""); // Blank line for counters
                var start = DateTime.Now;
                LogLine(0, $"Started at {start}");
                var explorer = Task.Run(Explorer);
                Thread.Sleep(500);
                var indexer = Task.Run(Indexer);
                Task.WaitAll(new Task[] { indexer, explorer });
                LogLine(0, consoleText);
                LogLine(0, $"Run took {DateTime.Now - start}");
                var collisions = $"Collision rate was {((indexerFailCount + explorerFailCount) / 2) / (double)Math.Min(indexerCount, explorerCount):P2}";
                LogLine(0, collisions);
                Console.WriteLine(collisions);
                Console.WriteLine("Done");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Unexpected Exception");
                Console.Error.WriteLine(ex.Message);
                return -1;
            }

            return 0;
        }

        static void LogLine(int level, string line)
        {
            if (level > logLevel)
                return;
            lock (logFile)
            {
                using (StreamWriter sw = new StreamWriter(logFile, true))
                {
                    sw.WriteLine(line);
                }
            }
        }

        static void Indexer()
        {
            string indexerFile = "";
            try
            {
                Guid IPropertyStoreGuid = new Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99");

                DirectoryInfo di = new DirectoryInfo(targetDirectory);

                bool done = false;
                while (!done)
                {
                    foreach (FileInfo file in di.GetFiles())
                    {
                        indexerFile = file.Name;
                        // To simulate indexer, open handler directly to get fine-grained control over flags
                        //HResult hr = (HResult)SHGetPropertyStoreFromParsingName(file.FullName, IntPtr.Zero,
                        //  GETPROPERTYSTOREFLAGS.GPS_DEFAULT, ref IPropertySto on '{indexerFile}'reGuid, out IPropertyStore ps);

                        var ps = new IPropertyHandler();
                        try
                        {
                            HResult hr = ps.Initialize(file.FullName, (uint)(StgmConstants.STGM_READ | StgmConstants.STGM_SHARE_DENY_NONE));

                            if (hr == HResult.Ok)
                            {
                                hr = ps.GetCount(out uint propertyCount);
                                if (hr == HResult.Ok)
                                {
                                    for (uint index = 0; index < propertyCount; index++)
                                    {
                                        PropVariant val = new PropVariant();
                                        hr = ps.GetAt(index, out PropertyKey key);
                                        if (hr == HResult.Ok)
                                        {
                                            hr = ps.GetValue(key, val);
                                            if (hr == HResult.Ok)
                                                LogLine(3, $"Indexer read {val}");
                                            else
                                                throw new Exception($"GetValue failed with 0x{hr:X}");
                                        }
                                        else
                                            throw new Exception($"GetAt failed with 0x{hr:X}");
                                    }
                                }
                                else if ((uint)hr == 0x80030021)
                                {
                                    LogLine(2, $"GetCount for {indexerFile} failed with STG_E_LOCKVIOLATION");
                                    IncrementIndexerFail();
                                }
                                else
                                    throw new Exception($"GetCount failed with 0x{hr:X}");

                                done = IncrementIndexer();
                                if (done)
                                    break;
                            }
                            else if ((uint)hr == 0x80030021)
                            {
                                LogLine(2, $"Open for {indexerFile} failed with STG_E_LOCKVIOLATION");
                                IncrementIndexerFail();
                            }
                            //else
                            //    throw new Exception($"Open failed with 0x{hr:X}");
                            else
                            {
                                // Indexer can tolerate some failures, but not too many
                                LogLine(1, $"Indexer Open failed on '{indexerFile}' with 0x{hr:X}");
                                IncrementIndexerFail();
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(ps);  // optional GC preempt
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogLine(1, $"Indexer terminated on '{indexerFile}': {ex}");
                IncrementIndexerTerminated();
            }
        }

        static void Explorer()
        {
            string explorerFile = "";
            try
            {
                Guid IPropertyStoreGuid = new Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99");

                PropertySystemNativeMethods.PSGetPropertyKeyFromName("System.Comment", out PropertyKey key);

                DirectoryInfo di = new DirectoryInfo(targetDirectory);
                int count = 0;
                bool done = false;
                while (!done)
                {
                    foreach (FileInfo file in di.GetFiles())
                    {
                        explorerFile = file.Name;
                        count++;
                        HResult hr = (HResult)SHGetPropertyStoreFromParsingName(file.FullName, IntPtr.Zero,
                        GETPROPERTYSTOREFLAGS.GPS_READWRITE, ref IPropertyStoreGuid, out IPropertyStore ps);

                        if (hr == HResult.Ok)
                        {
                            try
                            {
                                PropVariant val = new PropVariant();
                                hr = ps.GetValue(key, val);
                                if (hr == HResult.Ok)
                                {
                                    var sValue = val.ToString();

                                    if (count % 30 == 0)
                                    {
                                        PropVariant value = new PropVariant("test comment");
                                        hr = ps.SetValue(key, value);
                                        if (hr == HResult.Ok)
                                        {
                                            hr = ps.Commit();
                                            if (hr == HResult.Ok)
                                                LogLine(3, $"Explorer Read '{sValue}' and updated");
                                            else if ((int)hr > 0)
                                            {
                                                LogLine(2, $"Commit for {explorerFile} returned {hr:X}");
                                            }
                                            else
                                                throw new Exception($"Commit failed with 0x{hr:X}");
                                        }
                                        else
                                            throw new Exception($"SetValue failed with 0x{hr:X}");
                                    }
                                    else
                                        LogLine(3, $"Explorer Read '{sValue}'");
                                }
                                else if ((uint)hr == 0x80030021)
                                {
                                    LogLine(2, $"GetValue for {explorerFile} failed with STG_E_LOCKVIOLATION");
                                    IncrementExplorerFail();
                                }
                                else
                                    throw new Exception($"GetValue failed with 0x{hr:X}");

                                done = IncrementExplorer();
                                if (done)
                                    break;
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(ps);  // optional GC preempt
                            }
                        }
                        else if ((uint)hr == 0x80030021)
                        {
                            LogLine(2, $"Explorer Open failed on '{explorerFile}' with lock violation 0x{hr:X}");
                            IncrementExplorerFail();
                        }
                        else
                            throw new Exception($"Open failed with 0x{hr:X}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogLine(1, $"Explorer terminated for '{explorerFile}': {ex}");
                IncrementExplorerTerminated();
            }
        }

        private static bool IncrementIndexer() { return RewriteLine(ref indexerCount); }
        private static void IncrementIndexerFail() { RewriteLine(ref indexerFailCount); }
        private static void IncrementIndexerTerminated() { RewriteLine(ref indexerTerminatedCount); }
        private static bool IncrementExplorer() { return RewriteLine(ref explorerCount); }
        private static bool IncrementExplorerFail() { return RewriteLine(ref explorerFailCount); }
        private static bool IncrementExplorerTerminated() { return RewriteLine(ref explorerTerminatedCount); }
        private static long indexerCount = 0;
        private static long indexerFailCount = 0;
        private static long indexerTerminatedCount = 0;
        private static long explorerCount = 0;
        private static long explorerFailCount = 0;
        private static long explorerTerminatedCount = 0;
        private readonly static string consoleLock = "Console lock";
        private static string consoleText = "";
        private static bool RewriteLine(ref long count)
        {
            bool result = false;
            lock (consoleLock)
            {
                count++;
                consoleText = $"Indexer: {indexerCount} files, {indexerFailCount} failures, {indexerTerminatedCount} terminated. " +
                    $"Explorer: {explorerCount} files, {explorerFailCount} failures, {explorerTerminatedCount} terminated. ";
                int currentLineCursor = Console.CursorTop;
                Console.SetCursorPosition(0, currentLineCursor - 1);
                Console.Write(consoleText);
                Console.SetCursorPosition(0, currentLineCursor);
                result = (explorerCount > stopCount && indexerCount > stopCount) ||
                    explorerTerminatedCount > 0 || indexerTerminatedCount > 0;
            }
            return result;
        }

        public enum GETPROPERTYSTOREFLAGS
        {
            // If no flags are specified (GPS_DEFAULT), a read-only property store is returned that includes properties for the file or item.
            // In the case that the shell item is a file, the property store contains:
            //     1. properties about the file from the file system
            //     2. properties from the file itself provided by the file's property handler, unless that file is offline,
            //     see GPS_OPENSLOWITEM
            //     3. if requested by the file's property handler and supported by the file system, properties stored in the
            //     alternate property store.
            //
            // Non-file shell items should return a similar read-only store
            //
            // Specifying other GPS_ flags modifies the store that is returned
            GPS_DEFAULT = 0x00000000,
            GPS_HANDLERPROPERTIESONLY = 0x00000001,   // only include properties directly from the file's property handler
            GPS_READWRITE = 0x00000002,   // Writable stores will only include handler properties
            GPS_TEMPORARY = 0x00000004,   // A read/write store that only holds properties for the lifetime of the IShellItem object
            GPS_FASTPROPERTIESONLY = 0x00000008,   // do not include any properties from the file's property handler (because the file's property handler will hit the disk)
            GPS_OPENSLOWITEM = 0x00000010,   // include properties from a file's property handler, even if it means retrieving the file from offline storage.
            GPS_DELAYCREATION = 0x00000020,   // delay the creation of the file's property handler until those properties are read, written, or enumerated
            GPS_BESTEFFORT = 0x00000040,   // For readonly stores, succeed and return all available properties, even if one or more sources of properties fails. Not valid with GPS_READWRITE.
            GPS_NO_OPLOCK = 0x00000080,   // some data sources protect the read property store with an oplock, this disables that
            GPS_MASK_VALID = 0x000000FF,
        }
        [Flags]
        public enum StgmConstants
        {
            STGM_READ = 0x0,
            STGM_WRITE = 0x1,
            STGM_READWRITE = 0x2,
            STGM_SHARE_DENY_NONE = 0x40,
            STGM_SHARE_DENY_READ = 0x30,
            STGM_SHARE_DENY_WRITE = 0x20,
            STGM_SHARE_EXCLUSIVE = 0x10,
            STGM_PRIORITY = 0x40000,
            STGM_CREATE = 0x1000,
            STGM_CONVERT = 0x20000,
            STGM_FAILIFTHERE = 0x0,
            STGM_DIRECT = 0x0,
            STGM_TRANSACTED = 0x10000,
            STGM_NOSCRATCH = 0x100000,
            STGM_NOSNAPSHOT = 0x200000,
            STGM_SIMPLE = 0x8000000,
            STGM_DIRECT_SWMR = 0x400000,
            STGM_DELETEONRELEASE = 0x4000000
        }
        [DllImport("shell32.dll", SetLastError = true)]
        public static extern int SHGetPropertyStoreFromParsingName(
                [In][MarshalAs(UnmanagedType.LPWStr)] string pszPath,
                IntPtr zeroWorks,
                GETPROPERTYSTOREFLAGS flags,
                ref Guid iIdPropStore,
                [Out] out IPropertyStore propertyStore);


    }
    //PropertyHandler
    [ComImport]
    //#if x64
    [Guid("D06391EE-2FEB-419B-9667-AD160D0849F3")]
    //#else
    //  [Guid("60211757-EF87-465e-B6C1-B37CF98295F9")]
    //#endif
    class CPropertyHandlerClass
    {
    }

    [ComImport]
    [CoClass(typeof(CPropertyHandlerClass))]
    [Guid("B7D14566-0509-4CCE-A71F-0A554233BD9B")]
    interface IPropertyHandler : IInitializeWithFile, IPropertyStore, IPropertyStoreCapabilities
    {
    }

    [ComImport]
    [Guid("B7D14566-0509-4CCE-A71F-0A554233BD9B")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IInitializeWithFile
    {
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult Initialize(string pszFilePath, uint grfMode);
    }

    /// <summary>
    /// A property store
    /// </summary>
    [ComImport]
    [Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPropertyStore
    {
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetCount([Out] out uint propertyCount);

        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetAt([In] uint propertyIndex, out PropertyKey key);

        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetValue([In] ref PropertyKey key, [Out] PropVariant pv);

        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult SetValue([In] ref PropertyKey key, [In] PropVariant pv);

        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult Commit();
    }

    [ComImport]
    [Guid("C8E2D566-186E-4D49-BF41-6909EAD56ACC")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPropertyStoreCapabilities
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult IsPropertyWritable([In] ref PropertyKey propertyKey);
    }
}
