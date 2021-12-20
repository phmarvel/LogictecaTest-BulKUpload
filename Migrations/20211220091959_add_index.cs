using Microsoft.EntityFrameworkCore.Migrations;

namespace LogictecaTest.Migrations
{
    public partial class add_index : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AlterColumn<string>(
                name: "Part_SKU",
                table: "Items",
                type: "nvarchar(450)",
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(max)",
                oldNullable: true);

            migrationBuilder.CreateIndex(
                name: "IX_Items_Part_SKU",
                table: "Items",
                column: "Part_SKU",
                unique: true,
                filter: "[Part_SKU] IS NOT NULL");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropIndex(
                name: "IX_Items_Part_SKU",
                table: "Items");

            migrationBuilder.AlterColumn<string>(
                name: "Part_SKU",
                table: "Items",
                type: "nvarchar(max)",
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(450)",
                oldNullable: true);
        }
    }
}
