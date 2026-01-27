using System;
using System.Collections.Generic;
using System.Linq;
using HogarProtegido.Treasury.Models;

namespace HogarProtegido.Treasury.Services
{
    public class DataService
    {
        public DataService()
        {
            using var db = new TreasuryDbContext();
            db.Database.EnsureCreated();
        }

        public List<Movimiento> LoadMovimientos()
        {
            try
            {
                using var db = new TreasuryDbContext();
                return db.Movimientos.OrderByDescending(m => m.Fecha).ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al cargar: {ex.Message}");
                return new List<Movimiento>();
            }
        }

        public void SaveMovimientos(IEnumerable<Movimiento> movimientos)
        {
            try
            {
                using var db = new TreasuryDbContext();
                
                // Sincronización Senior: 
                // 1. Eliminar de la DB lo que ya no existe en la colección
                var incomingIds = movimientos.Select(m => m.Id).Where(id => id != Guid.Empty).ToList();
                var toDelete = db.Movimientos.AsEnumerable().Where(m => !incomingIds.Contains(m.Id)).ToList();
                
                if (toDelete.Any())
                {
                    db.Movimientos.RemoveRange(toDelete);
                }

                // 2. Upsert (Actualizar o Insertar)
                foreach (var m in movimientos)
                {
                    var existing = db.Movimientos.Find(m.Id);
                    if (existing != null)
                    {
                        db.Entry(existing).CurrentValues.SetValues(m);
                    }
                    else
                    {
                        db.Movimientos.Add(m);
                    }
                }
                
                db.SaveChanges();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error al sincronizar: {ex.Message}");
            }
        }

        public void AddMovimiento(Movimiento m)
        {
            using var db = new TreasuryDbContext();
            db.Movimientos.Add(m);
            db.SaveChanges();
        }
    }
}
