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
    
    public partial class Driver
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Driver()
        {
            this.DRV_Duty = new HashSet<DRV_Duty>();
            this.MAN_Data = new HashSet<MAN_Data>();
        }
    
        public int Id { get; set; }
        public Nullable<int> Employee_Number { get; set; }
        public string First_Name { get; set; }
        public string Second_Name { get; set; }
        public string Surename { get; set; }
        public string Type_Of_Employment { get; set; }
        public Nullable<double> Overtime_Rate { get; set; }
        public Nullable<double> Standard_Rate { get; set; }
        public string drv_card { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DRV_Duty> DRV_Duty { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<MAN_Data> MAN_Data { get; set; }
    }
}
