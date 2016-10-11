using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;

namespace FileMetadataAssociationManager
{
    // Sourced from this blog post http://social.msdn.microsoft.com/forums/en-US/wpf/thread/5909dbcc-9a9f-4260-bc36-de4aa9bbd383/ and copyright of its author
    public static class Extensions
    {
        public static int FindIndex<T>(this ObservableCollection<T> collection, Predicate<T> predicate)
        {
            int index = 0;
            foreach (var item in collection)
            {
                if (predicate(item))
                    return index;
                index++;
            }
            return -1;
        }

        public static IEnumerable<T> Union<T>(this IEnumerable<T> collection, IEnumerable<T> source, Func<T, T, bool> comparison)
        {
            IEnumerable<T> union = collection.Union(source, new LambdaComparer<T>(comparison));
            return union;
        }

        public static void Sort<T>(this ObservableCollection<T> collection, Comparison<T> comparison)
        {
            var comparer = new Comparer<T>(comparison);

            List<T> sorted = collection.OrderBy(x => x, comparer).ToList();

            for (int i = 0; i < sorted.Count(); i++)
                collection.Move(collection.IndexOf(sorted[i]), i);
        }

        private class Comparer<T> : IComparer<T>
        {
            private readonly Comparison<T> comparison;

            public Comparer(Comparison<T> comparison)
            {
                this.comparison = comparison;
            }

            #region IComparer<T> Members

            public int Compare(T x, T y)
            {
                return comparison.Invoke(x, y);
            }

            #endregion
        }
    }

    public class LambdaComparer<T> : IEqualityComparer<T>
    {
        private readonly Func<T, T, bool> _lambdaComparer;
        private readonly Func<T, int> _lambdaHash;

        public LambdaComparer(Func<T, T, bool> lambdaComparer) :
            this(lambdaComparer, o => 0)
        {
        }

        public LambdaComparer(Func<T, T, bool> lambdaComparer, Func<T, int> lambdaHash)
        {
            if (lambdaComparer == null)
                throw new ArgumentNullException("lambdaComparer");
            if (lambdaHash == null)
                throw new ArgumentNullException("lambdaHash");

            _lambdaComparer = lambdaComparer;
            _lambdaHash = lambdaHash;
        }

        public bool Equals(T x, T y)
        {
            return _lambdaComparer(x, y);
        }

        public int GetHashCode(T obj)
        {
            return _lambdaHash(obj);
        }
    }

    // This extension is necessary because RegistryKey.OpenBaseKey was not added until .Net framework 4
    // Code taken from http://stackoverflow.com/questions/26217199/what-are-some-alternatives-to-registrykey-openbasekey-in-net-3-5
    // No visible copyright
    public static class RegistryExtensions
    {

        public enum RegistryHiveType
        {
            X86,
            X64
        }

        static Dictionary<RegistryHive, UIntPtr> _hiveKeys = new Dictionary<RegistryHive, UIntPtr> {
        { RegistryHive.ClassesRoot, new UIntPtr(0x80000000u) },
        { RegistryHive.CurrentConfig, new UIntPtr(0x80000005u) },
        { RegistryHive.CurrentUser, new UIntPtr(0x80000001u) },
        { RegistryHive.DynData, new UIntPtr(0x80000006u) },
        { RegistryHive.LocalMachine, new UIntPtr(0x80000002u) },
        { RegistryHive.PerformanceData, new UIntPtr(0x80000004u) },
        { RegistryHive.Users, new UIntPtr(0x80000003u) }
    };

        static Dictionary<RegistryHiveType, RegistryAccessMask> _accessMasks = new Dictionary<RegistryHiveType, RegistryAccessMask> {
        { RegistryHiveType.X64, RegistryAccessMask.Wow6464 },
        { RegistryHiveType.X86, RegistryAccessMask.WoW6432 }
    };

        [Flags]
        public enum RegistryAccessMask
        {
            QueryValue = 0x0001,
            SetValue = 0x0002,
            CreateSubKey = 0x0004,
            EnumerateSubKeys = 0x0008,
            Notify = 0x0010,
            CreateLink = 0x0020,
            WoW6432 = 0x0200,
            Wow6464 = 0x0100,
            Write = 0x20006,
            Read = 0x20019,
            Execute = 0x20019,
            AllAccess = 0xF003F
        }

        [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
        public static extern int RegOpenKeyEx(
          UIntPtr hKey,
          string subKey,
          uint ulOptions,
          uint samDesired,
          out IntPtr hkResult);

        public static RegistryKey OpenBaseKey(RegistryHive registryHive, RegistryHiveType registryType)
        {
            UIntPtr hiveKey = _hiveKeys[registryHive];
            if (Environment.OSVersion.Platform == PlatformID.Win32NT && Environment.OSVersion.Version.Major > 5)
            {
                RegistryAccessMask flags = RegistryAccessMask.QueryValue | RegistryAccessMask.EnumerateSubKeys | RegistryAccessMask.SetValue | RegistryAccessMask.CreateSubKey | _accessMasks[registryType];
                IntPtr keyHandlePointer = IntPtr.Zero;
                int result = RegOpenKeyEx(hiveKey, String.Empty, 0, (uint)flags, out keyHandlePointer);
                if (result == 0)
                {
                    var safeRegistryHandleType = typeof(SafeHandleZeroOrMinusOneIsInvalid).Assembly.GetType("Microsoft.Win32.SafeHandles.SafeRegistryHandle");
                    var safeRegistryHandleConstructor = safeRegistryHandleType.GetConstructor(BindingFlags.Instance | BindingFlags.NonPublic, null, new[] { typeof(IntPtr), typeof(bool) }, null); // .NET < 4
                    if (safeRegistryHandleConstructor == null)
                        safeRegistryHandleConstructor = safeRegistryHandleType.GetConstructor(BindingFlags.Instance | BindingFlags.Public, null, new[] { typeof(IntPtr), typeof(bool) }, null); // .NET >= 4
                    var keyHandle = safeRegistryHandleConstructor.Invoke(new object[] { keyHandlePointer, true });
                    var net3Constructor = typeof(RegistryKey).GetConstructor(BindingFlags.Instance | BindingFlags.NonPublic, null, new[] { safeRegistryHandleType, typeof(bool) }, null);
                    var net4Constructor = typeof(RegistryKey).GetConstructor(BindingFlags.Instance | BindingFlags.NonPublic, null, new[] { typeof(IntPtr), typeof(bool), typeof(bool), typeof(bool), typeof(bool) }, null);
                    object key;
                    if (net4Constructor != null)
                        key = net4Constructor.Invoke(new object[] { keyHandlePointer, true, false, false, hiveKey == _hiveKeys[RegistryHive.PerformanceData] });
                    else if (net3Constructor != null)
                        key = net3Constructor.Invoke(new object[] { keyHandle, true });
                    else
                    {
                        var keyFromHandleMethod = typeof(RegistryKey).GetMethod("FromHandle", BindingFlags.Static | BindingFlags.Public, null, new[] { safeRegistryHandleType }, null);
                        key = keyFromHandleMethod.Invoke(null, new object[] { keyHandle });
                    }
                    var field = typeof(RegistryKey).GetField("keyName", BindingFlags.Instance | BindingFlags.NonPublic);
                    if (field != null)
                        field.SetValue(key, String.Empty);
                    return (RegistryKey)key;
                }
                else if (result == 2) // The key does not exist.
                    return null;
                throw new Win32Exception(result);
            }
            throw new PlatformNotSupportedException("The platform or operating system must be Windows 2000 or later.");
        }
    } 
}
