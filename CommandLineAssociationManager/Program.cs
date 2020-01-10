// Copyright (c) 2016, Dijii, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FileMetadataAssociationManager;
using NDesk.Options;
using AssociationMessages;

namespace CommandLineAssociationManager
{
    class Program
    {

         static string[] helpText = new string[] {
                "",
                "Usage:",
                "",
                "   FileMetaAssoc.exe {-l|-a|-r|-h} [-p=<profile name>] ",
                "                     [-d=<definition file name>] [<extension> ...]",
                "",
                "Where:",
                "",
                "   -l, --list",
                "      List all extensions with the File Meta property handler",
                "   -- OR --",
                "   -a, --add",
                "      Add the File Meta property handler for the specified extension(s)",
                "      (Requires administrator privileges)",
                "   -- OR --",
                "   -r, --remove",
                "      Remove the File Meta property handler from the specified extension(s)",
                "      (Requires administrator privileges)",
                "   -- OR --",
                "   -h, --help",
                "      Display this help",
                "",
                "   -p=<profile name>, --profile=<profile name>",
                "      Profile to be used for -add",
                "",
                "   --d=<definition file name>, --definitions=<definition file name>",
                "      Profile definitions file to be used for -add",
                "",
                "   -m, --merge",
                "      If set -add merges any existing settings into a new profile. ",
                "      Otherwise, profile settings are used",
                "",
                "   <extension>",
                "      One or more target extensions for -add or -remove, for example, .txt",
                "",
            };

        static int Main(string[] args)
        {

            int commands = 0;
            bool help = false;
            bool list = false;
            bool add = false;
            bool remove = false;
            string profile = null;
            string definitions = null;
            bool merge = false;

            try
            {
                var argParser = new OptionSet() {
                    { "l|list", v => {list = v != null; commands++;} },
                    { "a|add",  v => {add = v != null; commands++;} },
                    { "r|remove", v => {remove = v != null; commands++;} },
                    { "h|?|help", v => {help = v != null; commands++;} },
                    { "p|profile=", v => profile = v },
                    { "d|definitions=", v => definitions = v },
                    { "m|merge", v => merge = true },
                };
                List<string> extensions = argParser.Parse(args);

                if (commands != 1)
                {
                    throw new AssocMgrException
                    {
                        Description = LocalizedMessages.ExactlyOneOf,
                        Exception=null,
                        ErrorCode = WindowsErrorCodes.ERROR_INVALID_PARAMETER
                    };
                }

                if (help)
                {
                    foreach (var line in helpText)
                        Console.WriteLine(line);
                    return 0;
                }

                State state = new State();
                state.Populate(definitions);

                if (list)
                {
                    foreach (var ext in state.Extensions.Where(e => e.PropertyHandlerState == HandlerState.Ours).
                                 Concat(state.Extensions.Where(e => e.PropertyHandlerState == HandlerState.Chained)))
                        Console.WriteLine(String.Format("{0,-5}\t{1,-12}\t{2}", ext.Name, ext.Profile.Name, ext.PropertyHandlerDisplay));
                    return 0;
                }

                if (extensions.Count == 0)
                    throw new AssocMgrException
                    {
                        Description = LocalizedMessages.AtLeastOneExtension,
                        Exception = null,
                        ErrorCode = WindowsErrorCodes.ERROR_INVALID_PARAMETER
                    }; 

                foreach (var ext in extensions)
                {
                    Extension e = state.GetExtensionByName(ext);

                    if (remove)
                    {
                        if (e == null)
                            throw new AssocMgrException
                            {
                                Description = String.Format(LocalizedMessages.NotRegisteredExtension, ext),
                                Exception = null,
                                ErrorCode = WindowsErrorCodes.ERROR_INVALID_PARAMETER
                            };

                        if (!(e.PropertyHandlerState == HandlerState.Ours || 
                              e.PropertyHandlerState == HandlerState.Chained ||
                              e.PropertyHandlerState == HandlerState.ProfileOnly))
                            throw new AssocMgrException
                            {
                                Description = String.Format(LocalizedMessages.DoesNotHaveHandler, ext),
                                Exception = null,
                                ErrorCode = WindowsErrorCodes.ERROR_INVALID_PARAMETER
                            };

                        e.RemoveHandlerFromExtension();
                        Console.WriteLine(String.Format(LocalizedMessages.HandlerRemovedOK, e.Name));
                    }

                    if (add)
                    {

                        if (profile == null)
                            throw new AssocMgrException
                            {
                                Description = LocalizedMessages.AddNeedsProfile,
                                Exception = null,
                                ErrorCode = WindowsErrorCodes.ERROR_INVALID_PARAMETER
                            };

                        Profile p = state.GetProfileByName(profile);
                        if (p == null)
                            throw new AssocMgrException
                            {
                                Description = definitions != null ? String.Format(LocalizedMessages.NotAProfile, profile) : 
                                                                    String.Format(LocalizedMessages.NotAProfileMayNeedDefinitions, profile),
                                Exception = null,
                                ErrorCode = WindowsErrorCodes.ERROR_INVALID_PARAMETER
                            };

                        if (e == null)
                        {
                            e = state.CreateExtension(ext);
                        }

                        e.SetupHandlerForExtension(p, merge);
                        Console.WriteLine(String.Format(LocalizedMessages.HandlerAddedOK, e.Name));
                    }
                }
            }
            catch (AssocMgrException ae)
            {
                Console.Error.WriteLine(ae.DisplayString);
                return (int)ae.ErrorCode;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(LocalizedMessages.UnexpectedException);
                Console.Error.WriteLine(ex.Message);
                return -1;
            }

            return 0;
        }
    }
}
