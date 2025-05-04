namespace ExcelUpload.Models
{
    public class ExcelInfo
    {
        public int Id { get; set; }
        public string CardNumber { get; set; }
        public string CustomerName { get; set; }
        public decimal CustomerLimit { get; set; }
        public decimal CardBalance { get; set; }
        public decimal LoanBalance { get; set; }
        public decimal TotalBalance { get; set; }
        public decimal MinDue { get; set; }
        public int Buckets { get; set; }
        public int Pgs { get; set; }
        public string CustomerNumber { get; set; }
        public string InvoiceNumber { get; set; }
        public string SourceFile { get; set; }

    }
}
