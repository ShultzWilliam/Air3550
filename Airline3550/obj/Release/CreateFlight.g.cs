#pragma checksum "..\..\CreateFlight.xaml" "{8829d00f-11b8-4213-878b-770e8597ac16}" "2158240AFEE52297CBBA7C40829235BC04684E0DB93DE53A12E3C80CBF9187C4"
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
    /// CreateFlight
    /// </summary>
    public partial class CreateFlight : System.Windows.Controls.Page, System.Windows.Markup.IComponentConnector {
        
        
        #line 44 "..\..\CreateFlight.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox Origin;
        
        #line default
        #line hidden
        
        
        #line 60 "..\..\CreateFlight.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox Destination;
        
        #line default
        #line hidden
        
        
        #line 76 "..\..\CreateFlight.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox Departure_Date;
        
        #line default
        #line hidden
        
        
        #line 77 "..\..\CreateFlight.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox Departure_Time;
        
        #line default
        #line hidden
        
        
        #line 78 "..\..\CreateFlight.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox Departure_Terminal;
        
        #line default
        #line hidden
        
        
        #line 79 "..\..\CreateFlight.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox Arrival_Date;
        
        #line default
        #line hidden
        
        
        #line 80 "..\..\CreateFlight.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox Arrival_Time;
        
        #line default
        #line hidden
        
        
        #line 81 "..\..\CreateFlight.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox Arrival_Terminal;
        
        #line default
        #line hidden
        
        
        #line 82 "..\..\CreateFlight.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox Price;
        
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
            System.Uri resourceLocater = new System.Uri("/Air3550;component/createflight.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\CreateFlight.xaml"
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
            
            #line 12 "..\..\CreateFlight.xaml"
            ((System.Windows.Controls.Button)(target)).Click += new System.Windows.RoutedEventHandler(this.Sign_Out);
            
            #line default
            #line hidden
            return;
            case 2:
            
            #line 13 "..\..\CreateFlight.xaml"
            ((System.Windows.Controls.Button)(target)).Click += new System.Windows.RoutedEventHandler(this.Main_Menu);
            
            #line default
            #line hidden
            return;
            case 3:
            this.Origin = ((System.Windows.Controls.ComboBox)(target));
            return;
            case 4:
            this.Destination = ((System.Windows.Controls.ComboBox)(target));
            return;
            case 5:
            this.Departure_Date = ((System.Windows.Controls.TextBox)(target));
            return;
            case 6:
            this.Departure_Time = ((System.Windows.Controls.TextBox)(target));
            return;
            case 7:
            this.Departure_Terminal = ((System.Windows.Controls.TextBox)(target));
            return;
            case 8:
            this.Arrival_Date = ((System.Windows.Controls.TextBox)(target));
            return;
            case 9:
            this.Arrival_Time = ((System.Windows.Controls.TextBox)(target));
            return;
            case 10:
            this.Arrival_Terminal = ((System.Windows.Controls.TextBox)(target));
            return;
            case 11:
            this.Price = ((System.Windows.Controls.TextBox)(target));
            return;
            }
            this._contentLoaded = true;
        }
    }
}

