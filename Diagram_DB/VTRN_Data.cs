//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace WrkWebApp.Diagram_DB
{
    using System;
    using System.Collections.Generic;
    
    public partial class VTRN_Data
    {
        public int Id { get; set; }
        public Nullable<int> Veh { get; set; }
        public Nullable<decimal> Vtrn_Monies { get; set; }
        public Nullable<System.DateTime> Vtrn_Date_Driver { get; set; }
        public string Vtrn_Veh_Code { get; set; }
    
        public virtual Vehicle Vehicle { get; set; }
    }
}
