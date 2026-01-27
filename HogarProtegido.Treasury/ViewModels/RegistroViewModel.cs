using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using HogarProtegido.Treasury.Models;

namespace HogarProtegido.Treasury.ViewModels
{
    public class RegistroViewModel : ViewModelBase
    {
        private string _conceptoIngreso = string.Empty;
        private decimal? _montoIngreso;
        
        private string _conceptoEgreso = string.Empty;
        private decimal? _montoEgreso;

        private DateTime _fecha = DateTime.Today;
        private readonly MainViewModel _mainViewModel;
        private bool _isEditMode;

        public bool IsEditMode
        {
            get => _isEditMode;
            set => SetProperty(ref _isEditMode, value);
        }

        public DateTime Today => DateTime.Today;

        // Propiedades de Dashbaord (Sesión Temporal)
        public ObservableCollection<Movimiento> SessionIngresos { get; } = new();
        public ObservableCollection<Movimiento> SessionEgresos { get; } = new();
        
        private decimal _totalIngresos;
        public decimal TotalIngresos
        {
            get => _totalIngresos;
            set => SetProperty(ref _totalIngresos, value);
        }

        private decimal _totalGastos;
        public decimal TotalGastos
        {
            get => _totalGastos;
            set => SetProperty(ref _totalGastos, value);
        }

        private decimal _saldoDia;
        public decimal SaldoDia
        {
            get => _saldoDia;
            set => SetProperty(ref _saldoDia, value);
        }

        public decimal SaldoTotal => _mainViewModel.SaldoTotal;

        public string ConceptoIngreso
        {
            get => _conceptoIngreso;
            set => SetProperty(ref _conceptoIngreso, value);
        }

        public decimal? MontoIngreso
        {
            get => _montoIngreso;
            set => SetProperty(ref _montoIngreso, value);
        }

        public string ConceptoEgreso
        {
            get => _conceptoEgreso;
            set => SetProperty(ref _conceptoEgreso, value);
        }

        public decimal? MontoEgreso
        {
            get => _montoEgreso;
            set => SetProperty(ref _montoEgreso, value);
        }

        public DateTime Fecha
        {
            get => _fecha;
            set 
            { 
                if (SetProperty(ref _fecha, value))
                {
                    CargarMovimientosDeFecha();
                }
            }
        }

        public ICommand AgregarCommand { get; }
        public ICommand GuardarLibroCommand { get; }
        public ICommand EliminarMovimientoCommand { get; }

        public RegistroViewModel(MainViewModel mainViewModel)
        {
            _mainViewModel = mainViewModel;
            AgregarCommand = new RelayCommand(AgregarASesion);
            GuardarLibroCommand = new RelayCommand(ConfirmarYGuardar);
            EliminarMovimientoCommand = new RelayCommand(EliminarMovimiento);
            
            CargarMovimientosDeFecha();
        }

        private void CargarMovimientosDeFecha()
        {
            SessionIngresos.Clear();
            SessionEgresos.Clear();

            var existentes = _mainViewModel.Movimientos
                .Where(m => m.Fecha.Date == Fecha.Date)
                .ToList();

            foreach (var m in existentes)
            {
                var copia = new Movimiento 
                { 
                    Id = m.Id, // CRÍTICO: Preservar ID para sincronización correcta
                    Concepto = m.Concepto, 
                    Monto = m.Monto, 
                    Fecha = m.Fecha, 
                    Tipo = m.Tipo 
                };

                if (m.Tipo == TipoMovimiento.Ingreso) SessionIngresos.Add(copia);
                else SessionEgresos.Add(copia);
            }

            IsEditMode = existentes.Count > 0;
            RecalcularSessionTotals();
        }

        private void RecalcularSessionTotals()
        {
            decimal sumIng = 0;
            foreach (var m in SessionIngresos) sumIng += m.Monto;
            
            decimal sumEgr = 0;
            foreach (var m in SessionEgresos) sumEgr += m.Monto;

            TotalIngresos = sumIng;
            TotalGastos = sumEgr;
            SaldoDia = sumIng - sumEgr;
            OnPropertyChanged(nameof(SaldoTotal));
        }

        private void AgregarASesion(object? parameter)
        {
            string tipoStr = parameter?.ToString() ?? "Ingreso";
            TipoMovimiento tipo = (tipoStr == "Ingreso") ? TipoMovimiento.Ingreso : TipoMovimiento.Egreso;

            string concepto = tipo == TipoMovimiento.Ingreso ? ConceptoIngreso : ConceptoEgreso;
            decimal? monto = tipo == TipoMovimiento.Ingreso ? MontoIngreso : MontoEgreso;

            if (string.IsNullOrWhiteSpace(concepto) || !monto.HasValue || monto <= 0) return;

            var nuevo = new Movimiento
            {
                Concepto = concepto,
                Monto = monto.Value,
                Fecha = Fecha,
                Tipo = tipo
            };

            if (tipo == TipoMovimiento.Ingreso)
            {
                SessionIngresos.Add(nuevo);
                ConceptoIngreso = string.Empty;
                MontoIngreso = null;
            }
            else
            {
                SessionEgresos.Add(nuevo);
                ConceptoEgreso = string.Empty;
                MontoEgreso = null;
            }

            RecalcularSessionTotals();
        }

        private void EliminarMovimiento(object? parameter)
        {
            if (parameter is Movimiento mov)
            {
                if (mov.Tipo == TipoMovimiento.Ingreso) SessionIngresos.Remove(mov);
                else SessionEgresos.Remove(mov);

                RecalcularSessionTotals();
            }
        }

        private void ConfirmarYGuardar(object? parameter)
        {
            // Ya no bloqueamos si las listas están vacías, permitimos "vaciar" un día.
            
            string mensaje = IsEditMode 
                ? $"¿Desea ACTUALIZAR el registro del día {Fecha:dd/MM/yyyy}?\n\n(Si las listas están vacías, el registro anterior se eliminará)."
                : $"¿Desea confirmar el cierre de caja para el día {Fecha:dd/MM/yyyy}?";

            if (SessionIngresos.Count == 0 && SessionEgresos.Count == 0 && !IsEditMode)
            {
                System.Windows.MessageBox.Show("No hay movimientos para guardar.", "Aviso", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return;
            }

            var result = System.Windows.MessageBox.Show(mensaje, "Confirmar Cierre de Caja", 
                System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Question);

            if (result != System.Windows.MessageBoxResult.Yes) return;

            // Iniciar actualización masiva (bloquea el guardado por cada item)
            _mainViewModel.SetUpdating(true);

            try
            {
                // 1. Eliminar antiguos (si existen)
                var antiguos = _mainViewModel.Movimientos
                    .Where(m => m.Fecha.Date == Fecha.Date)
                    .ToList();

                foreach (var a in antiguos) _mainViewModel.Movimientos.Remove(a);

                // 2. Insertar nuevos
                foreach (var m in SessionIngresos) _mainViewModel.Movimientos.Add(m);
                foreach (var m in SessionEgresos) _mainViewModel.Movimientos.Add(m);
            }
            finally
            {
                // Finalizar actualización (ejecuta un único guardado atómico)
                _mainViewModel.SetUpdating(false);
            }

            System.Windows.MessageBox.Show("¡Cierre de caja guardado con éxito!", "Éxito", 
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);

            // 3. Limpiar y refrescar
            SessionIngresos.Clear();
            SessionEgresos.Clear();
            CargarMovimientosDeFecha();
        }
    }
}
