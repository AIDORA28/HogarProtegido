using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using HogarProtegido.Treasury.Models;
using HogarProtegido.Treasury.Services;

namespace HogarProtegido.Treasury.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        private readonly DataService _dataService;
        private decimal _saldoTotal;
        private object? _currentView;
        private bool _isUpdating; // Bandera para actualizaciones masivas
        public string Hoy => System.DateTime.Now.ToString("dd 'de' MMMM, yyyy", new System.Globalization.CultureInfo("es-ES"));

        public decimal SaldoTotal
        {
            get => _saldoTotal;
            set => SetProperty(ref _saldoTotal, value);
        }

        public object? CurrentView
        {
            get => _currentView;
            set => SetProperty(ref _currentView, value);
        }

        public ObservableCollection<Movimiento> Movimientos { get; set; }

        public ICommand ShowRegistroCommand { get; }
        public ICommand ShowReportesCommand { get; }
        public ICommand SalirCommand { get; }

        public MainViewModel()
        {
            _dataService = new Services.DataService();
            var savedMovimientos = _dataService.LoadMovimientos();
            
            Movimientos = new ObservableCollection<Movimiento>(savedMovimientos);
            Movimientos.CollectionChanged += OnMovimientosChanged;
            
            ShowRegistroCommand = new RelayCommand(_ => { CurrentView = new RegistroViewModel(this); });
            ShowReportesCommand = new RelayCommand(_ => { CurrentView = new ReportesViewModel(this); });
            SalirCommand = new RelayCommand(_ => { System.Windows.Application.Current.Shutdown(); });

            // Vista inicial
            CurrentView = new RegistroViewModel(this);
            
            RecalcularSaldos();
        }

        public void RecalcularSaldos()
        {
            var ingresos = Movimientos.Where(m => m.Tipo == TipoMovimiento.Ingreso).Sum(m => m.Monto);
            var egresos = Movimientos.Where(m => m.Tipo == TipoMovimiento.Egreso).Sum(m => m.Monto);
            SaldoTotal = ingresos - egresos;
        }


        private void OnMovimientosChanged(object? sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            RecalcularSaldos();
            if (!_isUpdating)
            {
                _dataService.SaveMovimientos(Movimientos);
            }
        }

        public void ManualSave()
        {
            _dataService.SaveMovimientos(Movimientos);
        }

        public void SetUpdating(bool value)
        {
            _isUpdating = value;
            if (!value) ManualSave();
        }
    }
}
