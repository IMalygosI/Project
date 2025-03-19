using System;
using System.Collections.Generic;
using GOST_Control.Models;
using Microsoft.EntityFrameworkCore;

namespace GOST_Control.Context;

public partial class DimaBaseContext : DbContext
{
    public DimaBaseContext()
    {
    }

    public DimaBaseContext(DbContextOptions<DimaBaseContext> options)
        : base(options)
    {
    }

    public virtual DbSet<Gost> Gosts { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see https://go.microsoft.com/fwlink/?LinkId=723263.
        => optionsBuilder.UseNpgsql("Host=89.110.53.87:5522;Database=dima_base;Username=dima;password=dima");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Gost>(entity =>
        {
            entity.HasKey(e => e.GostId).HasName("gost_pk");

            entity.ToTable("Gost", "Project");

            entity.Property(e => e.GostId)
                .UseIdentityAlwaysColumn()
                .HasColumnName("GostID");
            entity.Property(e => e.CheckPageNumbering).HasColumnType("character varying");
            entity.Property(e => e.FontName).HasMaxLength(255);
            entity.Property(e => e.Name).HasMaxLength(255);
            entity.Property(e => e.RequiredSections).HasColumnType("character varying");
            entity.Property(e => e.TextAlignment).HasColumnType("character varying");
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
