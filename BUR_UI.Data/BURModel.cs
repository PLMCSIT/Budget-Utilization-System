using System.Collections.Generic;

namespace BUR_UI.Entities
{
    public class BURModel
    {
        public BURModel()
        {
            this.Particulars = new List<Items>();
        }
        public string BURNumber { get; set; }
        public string Office { get; set; }
        public string OfficeCode { get; set; }
        public string Payee_Number { get; set; }
        public string Payee { get; set; }
        public int OfficeheadId { get; set; }
        public string OfficeheadName { get; set; }
        public string OfficeheadPos { get; set; }
        public string Description { get; set; }
        public string PRNumber { get; set; }
        public string BStaff_Number { get; set; }
        public string Staff { get; set; }
        public string Position { get; set; }
        public string BDHead_Number { get; set; }
        public string BDHead { get; set; }
        public string BDHead_Pos { get; set; }
        public string Date { get; set; }
        public List<Items> Particulars { get; set; }
    }
}
