// Copyright (c) 2016, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TestDriverAssoc
{
    class Common
    {
        public static bool SetupProfileForTestRun(State state)
        {
            try
            {
                // Backup any saved state file because some add operations update it
                var fiSavedState = GetDefaultSavedStateInfo();

                state.BackupFileInfo = null;
                state.WasNoState = !fiSavedState.Exists;
                if (fiSavedState.Exists)
                {
                    state.BackupFileInfo = new FileInfo(fiSavedState.Directory.FullName + @"\SavedState.xml.Backup");
                    fiSavedState.CopyTo(state.BackupFileInfo.FullName, true);
                }

                // Move in our own saved state so that results are predictable
                var fiOurs = new FileInfo(Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "SavedState.xml"));
                fiOurs.CopyTo(fiSavedState.FullName, true);

                return true;
            }
            catch (Exception ex)
            {
                state.AddReport(String.Format("Profile setup failed with unexpected exception '{0}'", ex.Message));
                return false;
            }
        }

        public static bool RestoreProfileAfterTestRun(State state)
        {
            try
            {
                var fiSavedState = GetDefaultSavedStateInfo();

                // Restore the saved state file
                if (state.BackupFileInfo != null)
                {
                    state.BackupFileInfo.CopyTo(fiSavedState.FullName, true);
                    state.BackupFileInfo.Delete();
                }
                else if (state.WasNoState)
                {
                    // Handle the case where a state file that did not exist before was created
                    if (fiSavedState.Exists)
                        fiSavedState.Delete();
                }

                return true;
            }
            catch (Exception ex)
            {
                state.AddReport(String.Format("Profile restoration failed with unexpected exception '{0}'", ex.Message));
                return false;
            }
        }

        public static void ClearProfile()
        {
            var fiSavedState = GetDefaultSavedStateInfo();
            if (fiSavedState.Exists)
                fiSavedState.Delete();
        }

        public static void ResetProfile()
        {
            // Move in our own saved state so that results are predictable
            var fiSavedState = GetDefaultSavedStateInfo();
            var fiOurs = new FileInfo(Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "SavedState.xml"));
            fiOurs.CopyTo(fiSavedState.FullName, true);
        }

        public static bool AreFilesIdentical(FileInfo a, FileInfo b)
        {
            if (a.Exists && b.Exists)
            {
                string af = File.ReadAllText(a.FullName);
                string bf = File.ReadAllText(b.FullName);
                return af == bf;
            }
            else if (!a.Exists && !b.Exists)
                return true;
            else
                return false;
        }

        public static FileInfo GetDefaultSavedStateInfo()
        {
            return new FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"FileMeta\SavedState.xml"));
        }
    }
}
