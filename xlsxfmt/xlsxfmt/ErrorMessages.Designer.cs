
namespace xlsxfmt {
    using System;
    

    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class ErrorMessages {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal ErrorMessages() {
        }
        

        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("xlsxfmt.ErrorMessages", typeof(ErrorMessages).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }

        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        internal static string FormatFileError {
            get {
                return ResourceManager.GetString("FormatFileError", resourceCulture);
            }
        }

        internal static string NoSheets {
            get {
                return ResourceManager.GetString("NoSheets", resourceCulture);
            }
        }
        
        internal static string NRE {
            get {
                return ResourceManager.GetString("NRE", resourceCulture);
            }
        }

        internal static string NullColumnSource {
            get {
                return ResourceManager.GetString("NullColumnSource", resourceCulture);
            }
        }
        

        internal static string NullSheetSource {
            get {
                return ResourceManager.GetString("NullSheetSource", resourceCulture);
            }
        }
    }
}
