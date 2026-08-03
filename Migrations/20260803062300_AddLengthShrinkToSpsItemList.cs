using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace VelastoProductionSystem.Migrations
{
    /// <inheritdoc />
    public partial class AddLengthShrinkToSpsItemList : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<float>(
                name: "LengthShrinkMax",
                table: "SpsItemLists",
                type: "real",
                nullable: true);

            migrationBuilder.AddColumn<float>(
                name: "LengthShrinkMin",
                table: "SpsItemLists",
                type: "real",
                nullable: true);

            migrationBuilder.AddColumn<float>(
                name: "LengthShrinkStd",
                table: "SpsItemLists",
                type: "real",
                nullable: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "LengthShrinkMax",
                table: "SpsItemLists");

            migrationBuilder.DropColumn(
                name: "LengthShrinkMin",
                table: "SpsItemLists");

            migrationBuilder.DropColumn(
                name: "LengthShrinkStd",
                table: "SpsItemLists");
        }
    }
}
