using System;

namespace HogarProtegido.Treasury.Models
{
    public enum TipoMovimiento
    {
        Ingreso,
        Egreso
    }

    public class Movimiento
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public DateTime Fecha { get; set; } = DateTime.Now;
        public string Concepto { get; set; } = string.Empty;
        public decimal Monto { get; set; }
        public TipoMovimiento Tipo { get; set; }

        // Propiedad calculada para visualización
        public string TipoLetra => Tipo == TipoMovimiento.Ingreso ? "💰" : "💸";
    }
}
