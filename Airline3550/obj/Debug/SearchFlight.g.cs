﻿#pragma checksum "..\..\SearchFlight.xaml" "{8829d00f-11b8-4213-878b-770e8597ac16}" "361E78041216916BB6D5A18B40D3F84D473E638E27DFAA2A077C857D4235497C"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using Airline3550;
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


namespace Air3550 {
    
    
    /// <summary>
    /// SearchFlight
    /// </summary>
    public partial class SearchFlight : System.Windows.Controls.Page, System.Windows.Markup.IComponentConnector {
        
        
        #line 28 "..\..\SearchFlight.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox Start;
        
        #line default
        #line hidden
        
        
        #line 44 "..\..\SearchFlight.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox End;
        
        #line default
        #line hidden
        
        
        #line 60 "..\..\SearchFlight.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.DatePicker Departure;
        
        #line default
        #line hidden
        
        
        #line 61 "..\..\SearchFlight.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.DatePicker Arrival;
        
        #line default
        #line hidden
        
        
        #line 63 "..\..\SearchFlight.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.CheckBox RoundTrip;
        
        #line default
        #line hidden
        
        
        #line 64 "..\..\SearchFlight.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.DataGrid Flights;
        
        #line default
        #line hidden
        
        
        #line 79 "..\..\SearchFlight.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox FlightID;
        
        #line default
        #line hidden
        
        
        #line 81 "..\..\SearchFlight.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBlock Warning;
        
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
            System.Uri resourceLocater = new System.Uri("/Air3550;component/searchflight.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\SearchFlight.xaml"
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
            
            #line 11 "..\..\SearchFlight.xaml"
            ((System.Windows.Controls.Button)(target)).Click += new System.Windows.RoutedEventHandler(this.Main_Menu);
            
            #line default
            #line hidden
            return;
            case 2:
            this.Start = ((System.Windows.Controls.ComboBox)(target));
            return;
            case 3:
            this.End = ((System.Windows.Controls.ComboBox)(target));
            return;
            case 4:
            this.Departure = ((System.Windows.Controls.DatePicker)(target));
            return;
            case 5:
            this.Arrival = ((System.Windows.Controls.DatePicker)(target));
            return;
            case 6:
            
            #line 62 "..\..\SearchFlight.xaml"
            ((System.Windows.Controls.Button)(target)).Click += new System.Windows.RoutedEventHandler(this.Search_Click);
            
            #line default
            #line hidden
            return;
            case 7:
            this.RoundTrip = ((System.Windows.Controls.CheckBox)(target));
            return;
            case 8:
            this.Flights = ((System.Windows.Controls.DataGrid)(target));
            return;
            case 9:
            
            #line 74 "..\..\SearchFlight.xaml"
            ((System.Windows.Controls.Button)(target)).Click += new System.Windows.RoutedEventHandler(this.Sign_Out);
            
            #line default
            #line hidden
            return;
            case 10:
            
            #line 78 "..\..\SearchFlight.xaml"
            ((System.Windows.Controls.Button)(target)).Click += new System.Windows.RoutedEventHandler(this.Submit_Click);
            
            #line default
            #line hidden
            return;
            case 11:
            this.FlightID = ((System.Windows.Controls.TextBox)(target));
            return;
            case 12:
            this.Warning = ((System.Windows.Controls.TextBlock)(target));
            return;
            }
            this._contentLoaded = true;
        }
    }
}

