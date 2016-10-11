﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace AssociationMessages {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    public class LocalizedMessages {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal LocalizedMessages() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        public static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("AssociationMessages.LocalizedMessages", typeof(LocalizedMessages).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        public static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to A profile must be specified for the -add command.
        /// </summary>
        public static string AddNeedsProfile {
            get {
                return ResourceManager.GetString("AddNeedsProfile", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Administrator privileges required to add or remove the File Meta property handler.
        /// </summary>
        public static string AdminPrivilegesNeeded {
            get {
                return ResourceManager.GetString("AdminPrivilegesNeeded", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to -add and -remove commands require at least one extension to be specified.
        /// </summary>
        public static string AtLeastOneExtension {
            get {
                return ResourceManager.GetString("AtLeastOneExtension", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Cannot add handler - our property handler is not registered.
        /// </summary>
        public static string CannotAddHandler {
            get {
                return ResourceManager.GetString("CannotAddHandler", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Cannot delete &apos;{0}&apos; as it is in use by at least one handler.
        /// </summary>
        public static string CannotDeleteProfile {
            get {
                return ResourceManager.GetString("CannotDeleteProfile", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Changing Profiles.
        /// </summary>
        public static string ChangingProfilesHeader {
            get {
                return ResourceManager.GetString("ChangingProfilesHeader", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to These changes will be applied to the following extensions: {0}. Do you want to proceed?.
        /// </summary>
        public static string ChangingProfilesQuestion {
            get {
                return ResourceManager.GetString("ChangingProfilesQuestion", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Closing with changes.
        /// </summary>
        public static string ClosingWithChanges {
            get {
                return ResourceManager.GetString("ClosingWithChanges", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to This will extend the existing handler, and create a custom profile for the extension using its existing properties plus the properties from the ‘{0}’ profile. Do you want to proceed?.
        /// </summary>
        public static string ConfirmCustomMerge {
            get {
                return ResourceManager.GetString("ConfirmCustomMerge", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to This will extend the existing handlers, and create a custom profile for each extension using its existing properties plus the properties from the ‘{0}’ profile. Do you want to proceed?.
        /// </summary>
        public static string ConfirmCustomMerges {
            get {
                return ResourceManager.GetString("ConfirmCustomMerges", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to This will extend the existing handler, and create a custom profile for the extension using its existing properties. Do you want to proceed?.
        /// </summary>
        public static string ConfirmCustomNoMerge {
            get {
                return ResourceManager.GetString("ConfirmCustomNoMerge", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to This will extend the existing handlers, and create a custom profile for each extension using its existing properties. Do you want to proceed?.
        /// </summary>
        public static string ConfirmCustomNoMerges {
            get {
                return ResourceManager.GetString("ConfirmCustomNoMerges", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Extension &apos;{0}&apos; does not currently have the File Meta property handler.
        /// </summary>
        public static string DoesNotHaveHandler {
            get {
                return ResourceManager.GetString("DoesNotHaveHandler", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Error.
        /// </summary>
        public static string ErrorHeader {
            get {
                return ResourceManager.GetString("ErrorHeader", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to You must specify exactly one of -list, -add, -remove or -help.
        /// </summary>
        public static string ExactlyOneOf {
            get {
                return ResourceManager.GetString("ExactlyOneOf", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Extension &apos;{0}&apos; already has the File Meta property handler.
        /// </summary>
        public static string ExtensionAlreadyHasHandler {
            get {
                return ResourceManager.GetString("ExtensionAlreadyHasHandler", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to File Meta property handler successfully added for extension &apos;{0}&apos;.
        /// </summary>
        public static string HandlerAddedOK {
            get {
                return ResourceManager.GetString("HandlerAddedOK", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to File Meta property handler successfully removed for extension &apos;{0}&apos;.
        /// </summary>
        public static string HandlerRemovedOK {
            get {
                return ResourceManager.GetString("HandlerRemovedOK", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Some errors occurred during handler setup.
        /// </summary>
        public static string HandlerSetupIssues {
            get {
                return ResourceManager.GetString("HandlerSetupIssues", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Help.
        /// </summary>
        public static string HelpHeader {
            get {
                return ResourceManager.GetString("HelpHeader", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Could not find definitions file &apos;{0}&apos;.
        /// </summary>
        public static string MissingDefinitionsFile {
            get {
                return ResourceManager.GetString("MissingDefinitionsFile", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to New Profile.
        /// </summary>
        public static string NewProfileName {
            get {
                return ResourceManager.GetString("NewProfileName", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to No context menus - our context handler is not registered.
        /// </summary>
        public static string NoContextMenus {
            get {
                return ResourceManager.GetString("NoContextMenus", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Administrator does not have write permission for &apos;{0}&apos; registry key.
        /// </summary>
        public static string NoRegistryPermission {
            get {
                return ResourceManager.GetString("NoRegistryPermission", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to No registry updates are required.
        /// </summary>
        public static string NoRegistryUpdatesNeeded {
            get {
                return ResourceManager.GetString("NoRegistryUpdatesNeeded", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to &apos;{0}&apos; is not the name of a defined profile.
        /// </summary>
        public static string NotAProfile {
            get {
                return ResourceManager.GetString("NotAProfile", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to &apos;{0}&apos; is not the name of a defined profile, you may need to specify a definitions file defining it.
        /// </summary>
        public static string NotAProfileMayNeedDefinitions {
            get {
                return ResourceManager.GetString("NotAProfileMayNeedDefinitions", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Extension &apos;{0}&apos; is not a valid registered extension.
        /// </summary>
        public static string NotRegisteredExtension {
            get {
                return ResourceManager.GetString("NotRegisteredExtension", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to &lt;None&gt;.
        /// </summary>
        public static string NullProfile {
            get {
                return ResourceManager.GetString("NullProfile", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Please select a profile other than &lt;None&gt;.
        /// </summary>
        public static string PleaseSelectProfile {
            get {
                return ResourceManager.GetString("PleaseSelectProfile", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to A profile called &apos;{0}&apos; already exists.
        /// </summary>
        public static string ProfileAlreadyExists {
            get {
                return ResourceManager.GetString("ProfileAlreadyExists", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Built in.
        /// </summary>
        public static string ProfileBuiltIn {
            get {
                return ResourceManager.GetString("ProfileBuiltIn", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Custom.
        /// </summary>
        public static string ProfileCustom {
            get {
                return ResourceManager.GetString("ProfileCustom", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Could not configure Profile.
        /// </summary>
        public static string ProfileError {
            get {
                return ResourceManager.GetString("ProfileError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Current property handler for {0} is {1}.
        /// </summary>
        public static string PropertyHandlerCurrent {
            get {
                return ResourceManager.GetString("PropertyHandlerCurrent", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to None.
        /// </summary>
        public static string PropertyHandlerNone {
            get {
                return ResourceManager.GetString("PropertyHandlerNone", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to {0}, {1}.
        /// </summary>
        public static string PropertyKeyFormatString {
            get {
                return ResourceManager.GetString("PropertyKeyFormatString", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Unable to initialize PropVariant..
        /// </summary>
        public static string PropVariantInitializationError {
            get {
                return ResourceManager.GetString("PropVariantInitializationError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Multi-dimensional SafeArrays not supported..
        /// </summary>
        public static string PropVariantMultiDimArray {
            get {
                return ResourceManager.GetString("PropVariantMultiDimArray", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to String argument cannot be null or empty..
        /// </summary>
        public static string PropVariantNullString {
            get {
                return ResourceManager.GetString("PropVariantNullString", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to This Value type is not supported..
        /// </summary>
        public static string PropVariantTypeNotSupported {
            get {
                return ResourceManager.GetString("PropVariantTypeNotSupported", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Cannot be cast to unsupported type..
        /// </summary>
        public static string PropVariantUnsupportedType {
            get {
                return ResourceManager.GetString("PropVariantUnsupportedType", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Update Registry.
        /// </summary>
        public static string RegistryUpdate {
            get {
                return ResourceManager.GetString("RegistryUpdate", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Updated registry settings for {0} extensions.
        /// </summary>
        public static string RegistryUpdatesMade {
            get {
                return ResourceManager.GetString("RegistryUpdatesMade", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Remove Property Group.
        /// </summary>
        public static string RemovePropertyGroupHeader {
            get {
                return ResourceManager.GetString("RemovePropertyGroupHeader", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Do you want to remove {0} and its contents from Full details , Preview Panel and Info Tip?.
        /// </summary>
        public static string RemovePropertyGroupQuestion {
            get {
                return ResourceManager.GetString("RemovePropertyGroupQuestion", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Remove Property.
        /// </summary>
        public static string RemovePropertyHeader {
            get {
                return ResourceManager.GetString("RemovePropertyHeader", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Do you want to remove {0} from Full details, Preview Panel and Info Tip?.
        /// </summary>
        public static string RemovePropertyQuestion {
            get {
                return ResourceManager.GetString("RemovePropertyQuestion", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to You have made changes that may require Explorer to be restarted to become effective. Do you want to restart Explorer now?.
        /// </summary>
        public static string RestartExplorerNow {
            get {
                return ResourceManager.GetString("RestartExplorerNow", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Handler addition.
        /// </summary>
        public static string SetupHandler {
            get {
                return ResourceManager.GetString("SetupHandler", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Unexpected exception:.
        /// </summary>
        public static string UnexpectedException {
            get {
                return ResourceManager.GetString("UnexpectedException", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Use drag and drop to populate Full details and Preview Panel, and reorder their contents.
        /// </summary>
        public static string UseDragDrop {
            get {
                return ResourceManager.GetString("UseDragDrop", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Error parsing XML store of saved custom profiles.
        /// </summary>
        public static string XmlParseError {
            get {
                return ResourceManager.GetString("XmlParseError", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Error updating XML store of saved custom profiles.
        /// </summary>
        public static string XmlWriteError {
            get {
                return ResourceManager.GetString("XmlWriteError", resourceCulture);
            }
        }
    }
}
