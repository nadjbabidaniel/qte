﻿#pragma checksum "..\..\CopyTimeEntry.xaml" "{406ea660-64cf-4c82-b6f0-42d48172a799}" "EE3F9D18235396B92A1B173E6CDBD84E"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using DevExpress.Core;
using DevExpress.Xpf.Core;
using DevExpress.Xpf.Core.ConditionalFormatting;
using DevExpress.Xpf.Core.DataSources;
using DevExpress.Xpf.Core.Serialization;
using DevExpress.Xpf.Core.ServerMode;
using DevExpress.Xpf.DXBinding;
using DevExpress.Xpf.Editors;
using DevExpress.Xpf.Editors.DataPager;
using DevExpress.Xpf.Editors.DateNavigator;
using DevExpress.Xpf.Editors.ExpressionEditor;
using DevExpress.Xpf.Editors.Filtering;
using DevExpress.Xpf.Editors.Flyout;
using DevExpress.Xpf.Editors.Popups;
using DevExpress.Xpf.Editors.Popups.Calendar;
using DevExpress.Xpf.Editors.RangeControl;
using DevExpress.Xpf.Editors.Settings;
using DevExpress.Xpf.Editors.Settings.Extension;
using DevExpress.Xpf.Editors.Validation;
using DevExpress.Xpf.Grid;
using DevExpress.Xpf.Grid.ConditionalFormatting;
using DevExpress.Xpf.Grid.LookUp;
using DevExpress.Xpf.Grid.TreeList;
using Infralution.Localization.Wpf;
using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;


namespace QuickTimeEntry {
    
    
    /// <summary>
    /// CopyTimeEntry
    /// </summary>
    public partial class CopyTimeEntry : System.Windows.Window, System.Windows.Markup.IComponentConnector {
        
        
        #line 23 "..\..\CopyTimeEntry.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Label lblInfo4;
        
        #line default
        #line hidden
        
        
        #line 24 "..\..\CopyTimeEntry.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Label lblInfo5;
        
        #line default
        #line hidden
        
        
        #line 26 "..\..\CopyTimeEntry.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Label lblInfo0;
        
        #line default
        #line hidden
        
        
        #line 27 "..\..\CopyTimeEntry.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Label lblInfo1;
        
        #line default
        #line hidden
        
        
        #line 28 "..\..\CopyTimeEntry.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Label lblInfo3;
        
        #line default
        #line hidden
        
        
        #line 30 "..\..\CopyTimeEntry.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal DevExpress.Xpf.Editors.DateEdit deDatum;
        
        #line default
        #line hidden
        
        
        #line 31 "..\..\CopyTimeEntry.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal DevExpress.Xpf.Editors.TextEdit teDauer;
        
        #line default
        #line hidden
        
        
        #line 32 "..\..\CopyTimeEntry.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal DevExpress.Xpf.Editors.TextEdit teZusatz;
        
        #line default
        #line hidden
        
        
        #line 34 "..\..\CopyTimeEntry.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button btnCopyTimeEntry;
        
        #line default
        #line hidden
        
        
        #line 36 "..\..\CopyTimeEntry.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Image btnCopyTimeEntryImage;
        
        #line default
        #line hidden
        
        
        #line 40 "..\..\CopyTimeEntry.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button btnAbort;
        
        #line default
        #line hidden
        
        
        #line 42 "..\..\CopyTimeEntry.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Image btnCancelTimeentryImage_1;
        
        #line default
        #line hidden
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Uri resourceLocater = new System.Uri("/QuickTimeEntry;component/copytimeentry.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\CopyTimeEntry.xaml"
            System.Windows.Application.LoadComponent(this, resourceLocater);
            
            #line default
            #line hidden
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        void System.Windows.Markup.IComponentConnector.Connect(int connectionId, object target) {
            switch (connectionId)
            {
            case 1:
            this.lblInfo4 = ((System.Windows.Controls.Label)(target));
            return;
            case 2:
            this.lblInfo5 = ((System.Windows.Controls.Label)(target));
            return;
            case 3:
            this.lblInfo0 = ((System.Windows.Controls.Label)(target));
            return;
            case 4:
            this.lblInfo1 = ((System.Windows.Controls.Label)(target));
            return;
            case 5:
            this.lblInfo3 = ((System.Windows.Controls.Label)(target));
            return;
            case 6:
            this.deDatum = ((DevExpress.Xpf.Editors.DateEdit)(target));
            return;
            case 7:
            this.teDauer = ((DevExpress.Xpf.Editors.TextEdit)(target));
            return;
            case 8:
            this.teZusatz = ((DevExpress.Xpf.Editors.TextEdit)(target));
            return;
            case 9:
            this.btnCopyTimeEntry = ((System.Windows.Controls.Button)(target));
            
            #line 34 "..\..\CopyTimeEntry.xaml"
            this.btnCopyTimeEntry.Click += new System.Windows.RoutedEventHandler(this.btnCopyTimeEntry_Click);
            
            #line default
            #line hidden
            return;
            case 10:
            this.btnCopyTimeEntryImage = ((System.Windows.Controls.Image)(target));
            return;
            case 11:
            this.btnAbort = ((System.Windows.Controls.Button)(target));
            
            #line 40 "..\..\CopyTimeEntry.xaml"
            this.btnAbort.Click += new System.Windows.RoutedEventHandler(this.btnAbort_Click);
            
            #line default
            #line hidden
            return;
            case 12:
            this.btnCancelTimeentryImage_1 = ((System.Windows.Controls.Image)(target));
            return;
            }
            this._contentLoaded = true;
        }
    }
}

