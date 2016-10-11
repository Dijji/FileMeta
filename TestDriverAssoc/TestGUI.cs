// Copyright (c) 2016, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Automation.Peers;

using AutomationExample;

namespace TestDriverAssoc
{
    class TestGUI
    {
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
                RegState.Wipe(Const.TestExt); // Ensure that we start clean
                state.AddReport(String.Format("#{0}: Add a handler with Simple profile for the minimal extension {1}, then remove it again", state.TestCounter, Const.TestExt));

                RegState.CreateMinimalExtension(Const.TestExt);
                OpenAssociationManager(state);
                ScrollAndSelectExtension(state.extensions, Const.TestExt);
                SelectProfile(state.profiles, "Simple");
                state.addButton.GetInvokePattern().Invoke();
                CloseAssociationManager(state); // Need to do this before verifying the registry state as it seems to flush the changes through
                pass = VerifyRegistryState(Const.TestExt, ref Const.V15BuiltIn);

                if (!pass)
                    state.AddReport(String.Format("Add did not produce the expected final registry state for {0}", Const.TestExt));
                else
                {
                    OpenAssociationManager(state);
                    ScrollAndSelectExtension(state.extensions, Const.TestExt);
                    state.removeButton.GetInvokePattern().Invoke();
                    CloseAssociationManager(state);
                    var clear = new RegState();
                    pass = VerifyRegistryState(Const.TestExt, ref clear);
                    if (!pass)
                        state.AddReport(String.Format("Remove did not produce the expected final registry state for {0}", Const.TestExt));
                }

                RegState.Wipe(Const.TestExt);

                state.AddReport(String.Format("#{0}: {1}", state.TestCounter++, pass ? "Passed" : "Failed"));
                overallPass &= pass;

                pass = true;

                state.AddReport(String.Format("#{0}: Extend a handler for {1} with Tiny profile, then remove it again", state.TestCounter, Const.TestExt));

                Const.V15UnExtendedTiny.Zap(Const.TestExt);
                OpenAssociationManager(state);
                ScrollAndSelectExtension(state.extensions, Const.TestExt);
                SelectProfile(state.profiles, "tiny");
                state.addButton.GetInvokePattern().Invoke();
                DismissDialog(state.mainWindow, "Handler addition", "Yes");
                CloseAssociationManager(state);
                pass = VerifyRegistryState(Const.TestExt, ref Const.V15ExtendedTiny);

                if (!pass)
                    state.AddReport(String.Format("Add did not produce the expected final registry state for {0}", Const.TestExt));
                else
                {
                    // Check that saved state is as expected
                    pass = VerifySavedState("SavedStatePlus.xml");
                    if (!pass)
                        state.AddReport(String.Format("Add did not produce the expected saved state file for {0}", Const.TestExt));
                    else
                    {
                        OpenAssociationManager(state);
                        ScrollAndSelectExtension(state.extensions, Const.TestExt);
                        state.removeButton.GetInvokePattern().Invoke();
                        CloseAssociationManager(state);
                        pass = VerifyRegistryState(Const.TestExt, ref Const.V15UnExtendedTiny);
                        if (!pass)
                            state.AddReport(String.Format("Remove did not produce the expected final registry state for {0}", Const.TestExt));
                    }
                }

                RegState.Wipe(Const.TestExt);

                state.AddReport(String.Format("#{0}: {1}", state.TestCounter++, pass ? "Passed" : "Failed"));
                overallPass &= pass;
                pass = true;

                state.AddReport(String.Format("#{0}: Set up a 1.3 handler for extension {1}, and verify that it is upgraded correctly", state.TestCounter, Const.TestExt));

                var result = MessageBox.Show(
                    "This test presses the Update Registry button, which updates all File Meta registry entries, as well as the test target. Are you happy to run this test?",
                    "Do you want to run this one?", MessageBoxButton.YesNo);

                if (result == MessageBoxResult.Yes)
                {
                    Const.V13BuiltIn.Zap(Const.TestExt);
                    OpenAssociationManager(state);
                    state.updateButton.GetInvokePattern().Invoke();
                    CloseAssociationManager(state); // Need to do this before verifying the registry state as it seems to flush the changes through
                    pass = VerifyRegistryState(Const.TestExt, ref Const.V15BuiltIn);

                    if (!pass)
                        state.AddReport(String.Format("Upgrade did not produce the expected final registry state for {0}", Const.TestExt));

                    RegState.Wipe(Const.TestExt);

                    state.AddReport(String.Format("#{0}: {1}", state.TestCounter++, pass ? "Passed" : "Failed"));
                    overallPass &= pass;
                }
                else
                    state.AddReport(String.Format("#{0}: {1}", state.TestCounter++, "Skipped"));

                RegState.Wipe(Const.TestExt);
            }
            catch (Exception ex)
            {
                state.AddReport(String.Format("Run failed with unexpected exception '{0}'", ex.Message));
            }
            finally
            {
                RegState.Wipe(Const.TestExt); 
                Common.RestoreProfileAfterTestRun(state);
            }

            state.AddReport("");
            state.AddReport(String.Format("Run {0} completed {1}", state.RunCounter++, overallPass ? "with no failures" : "with some failures"));
        }

        private static bool VerifyRegistryState(string ext, ref RegState expected)
        {
            var outcome = new RegState();
            outcome.Read(Const.TestExt);

            // Verify that we got the final state
            return (outcome == expected);
        }

        public static bool VerifySavedState(string expected)
        {
            return Common.AreFilesIdentical(Common.GetDefaultSavedStateInfo(), new FileInfo(Path.Combine(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath), expected)));
        }

        private static void OpenAssociationManager(State state)
        {
            state.process = Process.Start( @"C:\Program Files\File Metadata\FileMetaAssociationManager.exe");

            state.process.WaitForInputIdle();
            Delay();

            state.mainWindow = AutomationElement.RootElement.FindChildByProcessId(state.process.Id);

            Delay();
            state.extensions = state.mainWindow.FindChildById("listExtensions");
            state.profiles = state.mainWindow.FindChildById("comboProfile");
            state.addButton = state.mainWindow.FindChildById("addHandler");
            state.removeButton = state.mainWindow.FindChildById("removeHandler");
            state.updateButton = state.mainWindow.FindChildById("updateRegistry");
        }

        private static void CloseAssociationManager(State state)
        {
            if (state.process != null)
            {
                state.process.CloseMainWindow();
                DismissDialog(state.mainWindow, "Closing with changes", "No");  // Needed In case we have made changes
                state.process.Dispose();
                state.process = null;
            }
        } 

        private static void SelectProfile(AutomationElement profiles, string profile)
        {
            // Ensure all profiles are displayed
            profiles.GetExpandCollapsePattern().Expand();
            //profiles.GetExpandCollapsePattern().Collapse();

            // There is a gotcha where TextBoxes in certain other controls are not visible except by walking the Raw View
            var child = TreeWalker.RawViewWalker.GetFirstChild(profiles);
            while (child != null)
            {
                var gc = TreeWalker.RawViewWalker.GetFirstChild(child);
                if (gc != null && gc.Current.ControlType == ControlType.Text)
                {
                    if (gc.Current.Name == profile)
                    {
                        child.GetSelectionItemPattern().Select();
                        return;
                    }
                }
                child = TreeWalker.RawViewWalker.GetNextSibling(child);
            }

            throw new Exception(String.Format("Failed to select profile {0}", profile));
        }

        private static void ScrollAndSelectExtension(AutomationElement extensions, string ext)
        {
            // There is a gotcha where TextBoxes in certain other controls are not visible except by walking the Raw View
            var s = extensions.GetScrollPattern();
            while (s.Current.VerticalScrollPercent > 0)
                s.ScrollVertical(ScrollAmount.LargeDecrement);

            while (true)
            {
                if (SelectExtension(extensions, ext))
                    return;

                if (s.Current.VerticalScrollPercent < 100)
                    s.ScrollVertical(ScrollAmount.LargeIncrement);
                else
                    break;
            }

            throw new Exception (String.Format("Failed to select extension {0}", ext));
        }

        private static bool SelectExtension(AutomationElement extensions, string ext)
        {
            // There is a gotcha where TextBoxes in certain other controls are not visible except by walking the Raw View
            var child = TreeWalker.RawViewWalker.GetFirstChild(extensions);
            while (child != null)
            {
                var gc = TreeWalker.RawViewWalker.GetFirstChild(child);
                if (gc != null && gc.Current.ControlType == ControlType.Custom)
                {
                    var ggc = TreeWalker.RawViewWalker.GetFirstChild(gc);
                    if (ggc != null && ggc.Current.ControlType == ControlType.Text)
                    {
                        if (String.Compare(ggc.Current.Name, ext, true) == 0)
                        {
                            child.GetSelectionItemPattern().Select(); 
                            return true;
                        }
                    }
                }
                child = TreeWalker.RawViewWalker.GetNextSibling(child);
            }
            return false;
        }

        private static void DismissDialog(AutomationElement mainWindow, string title, string buttonName)
        {
            Delay();

            var dialog = mainWindow.FindChildByName(title);
            if (dialog == null)
                dialog = AutomationElement.RootElement.FindChildByName(title);

            if (dialog != null)
            {
                var dlgButton = dialog.FindChildByName(buttonName);
                dlgButton.GetInvokePattern().Invoke();
            }
        }

        private static void Delay(int milliseconds = 1000)
        {
            Thread.Sleep(milliseconds);
        }
    }
}
