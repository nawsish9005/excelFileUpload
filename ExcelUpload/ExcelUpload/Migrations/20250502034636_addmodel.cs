using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace ExcelUpload.Migrations
{
    /// <inheritdoc />
    public partial class addmodel : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "ExcelInfos",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    CardNumber = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    CustomerName = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    CustomerLimit = table.Column<decimal>(type: "decimal(18,2)", nullable: false),
                    CardBalance = table.Column<decimal>(type: "decimal(18,2)", nullable: false),
                    LoanBalance = table.Column<decimal>(type: "decimal(18,2)", nullable: false),
                    TotalBalance = table.Column<decimal>(type: "decimal(18,2)", nullable: false),
                    MinDue = table.Column<decimal>(type: "decimal(18,2)", nullable: false),
                    Buckets = table.Column<int>(type: "int", nullable: false),
                    Pgs = table.Column<int>(type: "int", nullable: false),
                    CustomerNumber = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    InvoiceNumber = table.Column<string>(type: "nvarchar(max)", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ExcelInfos", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "ExcelInfos");
        }
    }
}
