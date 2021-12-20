﻿// <auto-generated />
using LogictecaTest.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Metadata;
using Microsoft.EntityFrameworkCore.Migrations;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;

namespace LogictecaTest.Migrations
{
    [DbContext(typeof(ApplicationDbContext))]
    [Migration("20211220091959_add_index")]
    partial class add_index
    {
        protected override void BuildTargetModel(ModelBuilder modelBuilder)
        {
#pragma warning disable 612, 618
            modelBuilder
                .HasAnnotation("Relational:MaxIdentifierLength", 128)
                .HasAnnotation("ProductVersion", "5.0.13")
                .HasAnnotation("SqlServer:ValueGenerationStrategy", SqlServerValueGenerationStrategy.IdentityColumn);

            modelBuilder.Entity("LogictecaTest.Models.Item", b =>
                {
                    b.Property<int>("Id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int")
                        .HasAnnotation("SqlServer:ValueGenerationStrategy", SqlServerValueGenerationStrategy.IdentityColumn);

                    b.Property<string>("Band")
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("Category_Code")
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("Discounted_Price")
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("Item_Description")
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("List_Price")
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("Manufacturer")
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("Minimum_Discount")
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("Part_SKU")
                        .HasColumnType("nvarchar(450)");

                    b.HasKey("Id");

                    b.HasIndex("Part_SKU")
                        .IsUnique()
                        .HasFilter("[Part_SKU] IS NOT NULL");

                    b.ToTable("Items");
                });
#pragma warning restore 612, 618
        }
    }
}
