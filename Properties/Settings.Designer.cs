﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.18444
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace LxWorkList.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "12.0.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.ConnectionString)]
        [global::System.Configuration.DefaultSettingValueAttribute("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\"\\\\OXNETFS02\\RenalDir\\Renal\\Departme" +
            "ntal\\Transplant Immunology\\TTSSPD\\AIFILES\\SolidOrgan.mdb\"")]
        public string SolidOrganDBLive {
            get {
                return ((string)(this["SolidOrganDBLive"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.ConnectionString)]
        [global::System.Configuration.DefaultSettingValueAttribute("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Temp\\SolidOrgan.mdb")]
        public string SolidOrganDBTest {
            get {
                return ((string)(this["SolidOrganDBTest"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("\\\\OXNETFS02\\RenalDir\\Renal\\Departmental\\Transplant Immunology\\TTSSPD\\Luminex\\Pati" +
            "ent lists\\")]
        public string LiveWorkListPath {
            get {
                return ((string)(this["LiveWorkListPath"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.ConnectionString)]
        [global::System.Configuration.DefaultSettingValueAttribute("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\\\OXNETFS02\\RenalDir\\Renal\\Departme" +
            "ntal\\Transplant Immunology\\FoxPro\\;Extended Properties=dBASE IV;User ID=Admin;")]
        public string RenalSytemDBF {
            get {
                return ((string)(this["RenalSytemDBF"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.ConnectionString)]
        [global::System.Configuration.DefaultSettingValueAttribute("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\\temp\\renal;User ID=Admin;Extended" +
            " Properties=\"dBASE IV\"")]
        public string RenalSystemDBFlocal {
            get {
                return ((string)(this["RenalSystemDBFlocal"]));
            }
        }
    }
}
