using Microsoft.EntityFrameworkCore;
using HogarProtegido.Treasury.Models;
using System.IO;
using System;

namespace HogarProtegido.Treasury.Services
{
    public class TreasuryDbContext : DbContext
    {
        public DbSet<Movimiento> Movimientos { get; set; }

        private readonly string _dbPath;

        public TreasuryDbContext()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folderPath = Path.Combine(appData, "HogarProtegidoTreasury");
            if (!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath);
            _dbPath = Path.Combine(folderPath, "tesoreria.db");
        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlite($"Data Source={_dbPath}");
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Movimiento>().HasKey(m => m.Id);
            modelBuilder.Entity<Movimiento>().Property(m => m.Concepto).IsRequired();
            modelBuilder.Entity<Movimiento>().Property(m => m.Monto).HasColumnType("decimal(18,2)");
        }
    }
}
