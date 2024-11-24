using Microsoft.EntityFrameworkCore.Migrations;

namespace crowlerSj.Migrations
{
    public partial class djsfgdeee : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.InsertData(
                table: "Settings",
                columns: new[] { "Id", "IsCrowl" },
                values: new object[] { 1L, false });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DeleteData(
                table: "Settings",
                keyColumn: "Id",
                keyValue: 1L);
        }
    }
}
