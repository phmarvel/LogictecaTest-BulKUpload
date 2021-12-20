using Microsoft.EntityFrameworkCore.Migrations;

namespace LogictecaTest.Migrations
{
    public partial class init_db : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Items",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Band = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Category_Code = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Manufacturer = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Part_SKU = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Item_Description = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    List_Price = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Minimum_Discount = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Discounted_Price = table.Column<string>(type: "nvarchar(max)", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Items", x => x.Id);
                });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Items");
        }
    }
}
