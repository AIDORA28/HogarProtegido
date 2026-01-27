using System;
using System.Collections.ObjectModel;
using System.Linq;
using HogarProtegido.Treasury.Models;

namespace HogarProtegido.Treasury.ViewModels
{
    public class DailyReportItem : ViewModelBase
    {
        public DateTime Fecha { get; set; }
        public ObservableCollection<Movimiento> Ingresos { get; set; } = new();
        public ObservableCollection<Movimiento> Egresos { get; set; } = new();
        
        public decimal TotalIngresos => Ingresos.Sum(m => m.Monto);
        public decimal TotalEgresos => Egresos.Sum(m => m.Monto);
        public decimal TotalDia => TotalIngresos - TotalEgresos;
        public bool IsTotalDiaNegativo => TotalDia < 0;
        
        private decimal _saldoAyer;
        public decimal SaldoAyer
        {
            get => _saldoAyer;
            set { SetProperty(ref _saldoAyer, value); OnPropertyChanged(nameof(SaldoHoy)); }
        }

        public decimal SaldoHoy => TotalDia + SaldoAyer;
    }


    public class ReportesViewModel : ViewModelBase
    {
        private readonly MainViewModel _mainViewModel;
        private int _selectedViewIndex = 0; // 0: Detallado, 1: Listado
        private DateTime? _fechaInicio = null;
        private DateTime? _fechaFin = null;
        private bool _filterByDateEnabled = false;

        public DateTime Today => DateTime.Today;

        public ObservableCollection<DailyReportItem> DailyReports { get; set; } = new();

        public DateTime? FechaInicio
        {
            get => _fechaInicio;
            set 
            { 
                if (SetProperty(ref _fechaInicio, value))
                {
                    if (_fechaInicio.HasValue && _fechaFin.HasValue && _fechaFin < _fechaInicio)
                    {
                        FechaFin = _fechaInicio;
                    }
                    FilterByDateEnabled = value != null || FechaFin != null; 
                }
            }
        }

        public DateTime? FechaFin
        {
            get => _fechaFin;
            set 
            { 
                if (SetProperty(ref _fechaFin, value))
                {
                    if (_fechaInicio.HasValue && _fechaFin.HasValue && _fechaFin < _fechaInicio)
                    {
                        FechaFin = _fechaInicio;
                    }
                    FilterByDateEnabled = value != null || FechaInicio != null; 
                }
            }
        }

        public bool FilterByDateEnabled
        {
            get => _filterByDateEnabled;
            set { SetProperty(ref _filterByDateEnabled, value); GenerarReporte(); }
        }

        private decimal _totalIngresosHistorico;
        public decimal TotalIngresosHistorico
        {
            get => _totalIngresosHistorico;
            set => SetProperty(ref _totalIngresosHistorico, value);
        }

        private decimal _totalEgresosHistorico;
        public decimal TotalEgresosHistorico
        {
            get => _totalEgresosHistorico;
            set => SetProperty(ref _totalEgresosHistorico, value);
        }


        public int SelectedViewIndex
        {
            get => _selectedViewIndex;
            set => SetProperty(ref _selectedViewIndex, value);
        }

        public System.Windows.Input.ICommand SelectViewCommand { get; }
        public System.Windows.Input.ICommand LimpiarFiltroCommand { get; }
        public System.Windows.Input.ICommand ExportarPDFCommand { get; }
        public System.Windows.Input.ICommand ExportarExcelCommand { get; }
        public System.Windows.Input.ICommand RestablecerCommand { get; }

        public ReportesViewModel(MainViewModel mainViewModel)
        {
            _mainViewModel = mainViewModel;
            SelectViewCommand = new RelayCommand(p => SelectedViewIndex = int.Parse(p?.ToString() ?? "0"));
            LimpiarFiltroCommand = new RelayCommand(_ => FilterByDateEnabled = false);
            ExportarPDFCommand = new RelayCommand(_ => ExportarAPDF());
            ExportarExcelCommand = new RelayCommand(_ => ExportarAExcel());
            RestablecerCommand = new RelayCommand(_ => { 
                _fechaInicio = null; 
                _fechaFin = null; 
                OnPropertyChanged(nameof(FechaInicio)); 
                OnPropertyChanged(nameof(FechaFin));
                FilterByDateEnabled = false; 
            });
            GenerarReporte();
        }

        private void ExportarAPDF()
        {
            if (DailyReports.Count == 0) return;

            var sfd = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Archivos PDF (*.pdf)|*.pdf",
                FileName = $"Reporte_Detallado_{DateTime.Now:yyyyMMdd_HHmm}.pdf"
            };

            if (sfd.ShowDialog() == true)
            {
                try
                {
                    using var writer = new iText.Kernel.Pdf.PdfWriter(sfd.FileName);
                    using var pdf = new iText.Kernel.Pdf.PdfDocument(writer);
                    var document = new iText.Layout.Document(pdf, iText.Kernel.Geom.PageSize.A4);
                    
                    document.SetMargins(40, 30, 50, 30); // Top, Right, Bottom, Left (más espacio para footer)

                    // Paleta de colores premium
                    var primaryColor = new iText.Kernel.Colors.DeviceRgb(0, 121, 107);      // Teal 700
                    var accentColor = new iText.Kernel.Colors.DeviceRgb(0, 150, 136);     // Teal 600
                    var successBg = new iText.Kernel.Colors.DeviceRgb(220, 237, 200);     // Light Green
                    var errorBg = new iText.Kernel.Colors.DeviceRgb(255, 205, 210);       // Light Red
                    var neutralBg = new iText.Kernel.Colors.DeviceRgb(245, 245, 245);     // Gray 50

                    // ===== CABECERA PRINCIPAL =====
                    var titleMain = new iText.Layout.Element.Paragraph("HOGAR PROTEGIDO")
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                        .SetFontSize(24)
                        .SetFontColor(primaryColor)
                        .SetMarginBottom(5);
                    document.Add(titleMain);

                    var subtitle = new iText.Layout.Element.Paragraph("REPORTE DETALLADO DE CAJA")
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                        .SetFontSize(14)
                        .SetFontColor(new iText.Kernel.Colors.DeviceRgb(100, 100, 100))
                        .SetMarginBottom(15);
                    document.Add(subtitle);

                    // Rango de fechas
                    var rangoText = FilterByDateEnabled 
                        ? $"Período: {FechaInicio:dd/MM/yyyy} - {FechaFin:dd/MM/yyyy}" 
                        : "Historial Completo de Movimientos";
                    var rangoPara = new iText.Layout.Element.Paragraph(rangoText)
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                        .SetFontSize(10)
                        .SetMarginBottom(20);
                    document.Add(rangoPara);

                    // ===== RESUMEN EJECUTIVO =====
                    var totalIngresos = DailyReports.Sum(d => d.TotalIngresos);
                    var totalEgresos = DailyReports.Sum(d => d.TotalEgresos);
                    var saldoFinal = DailyReports.LastOrDefault()?.SaldoHoy ?? 0;
                    var diasReportados = DailyReports.Count;

                    var summaryTable = new iText.Layout.Element.Table(iText.Layout.Properties.UnitValue.CreatePercentArray(new float[] { 25, 25, 25, 25 }))
                        .UseAllAvailableWidth()
                        .SetMarginBottom(25);

                    // Header del resumen
                    summaryTable.AddCell(new iText.Layout.Element.Cell(1, 4)
                        .Add(new iText.Layout.Element.Paragraph("RESUMEN EJECUTIVO")
                     .SetFontSize(11)
                            .SetFontColor(iText.Kernel.Colors.DeviceRgb.WHITE))
                        .SetBackgroundColor(primaryColor)
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                        .SetPadding(8));

                    // Datos del resumen
                    summaryTable.AddCell(CreateSummaryCell("Días Reportados", diasReportados.ToString(), neutralBg));
                    summaryTable.AddCell(CreateSummaryCell("Total Ingresos", $"S/ {totalIngresos:N2}", successBg));
                    summaryTable.AddCell(CreateSummaryCell("Total Egresos", $"S/ {totalEgresos:N2}", errorBg));
                    summaryTable.AddCell(CreateSummaryCell("Saldo Final", $"S/ {saldoFinal:N2}", new iText.Kernel.Colors.DeviceRgb(178, 223, 219)));

                    document.Add(summaryTable);

                    // Separador
                    document.Add(new iText.Layout.Element.LineSeparator(new iText.Kernel.Pdf.Canvas.Draw.SolidLine(1.5f))
                        .SetMarginBottom(15));

                    // ===== DETALLE DIARIO CON LÓGICA ANTI-PARTICIÓN =====
                    const float pageHeight = 842f; // A4 height in points
                    const float topMargin = 40f;
                    const float bottomMargin = 50f;
                    const float usableHeight = pageHeight - topMargin - bottomMargin;
                    float currentY = usableHeight - 250f; // Espacio ya usado por cabecera y resumen

                    foreach (var day in DailyReports)
                    {
                        // Estimar altura necesaria para este día (aproximado)
                        float estimatedHeightForDay = 50 + // Header del día
                            (day.Ingresos.Any() ? (35 + day.Ingresos.Count() * 15 + 20) : 50) + // Tabla ingresos
                            (day.Egresos.Any() ? (35 + day.Egresos.Count() * 15 + 20) : 50) +    // Tabla egresos
                            80; // Balance de cierre

                        // Si no cabe en la página actual, nueva página
                        if (currentY < estimatedHeightForDay)
                        {
                            document.Add(new iText.Layout.Element.AreaBreak(iText.Layout.Properties.AreaBreakType.NEXT_PAGE));
                            currentY = usableHeight;
                        }

                        var dayCard = new iText.Layout.Element.Div()
                            .SetBackgroundColor(new iText.Kernel.Colors.DeviceRgb(250, 250, 250))
                            .SetBorder(new iText.Layout.Borders.SolidBorder(accentColor, 1f))
                            .SetPadding(12)
                            .SetMarginBottom(15)
                            .SetKeepTogether(true);

                        // Encabezado de la fecha
                        var dateHeader = new iText.Layout.Element.Paragraph($"📅 {day.Fecha:dddd, dd 'de' MMMM 'de' yyyy}".ToUpper())
                            .SetFontSize(12)
                            .SetFontColor(primaryColor)
                            .SetMarginBottom(10);
                        dayCard.Add(dateHeader);

                        // === TABLA DE INGRESOS ===
                        var ingLabel = new iText.Layout.Element.Paragraph("💰 INGRESOS (+)")
                            .SetFontSize(10)
                            .SetFontColor(new iText.Kernel.Colors.DeviceRgb(46, 125, 50))
                            .SetMarginBottom(5);
                        dayCard.Add(ingLabel);

                        var tIng = new iText.Layout.Element.Table(iText.Layout.Properties.UnitValue.CreatePercentArray(new float[] { 75, 25 }))
                            .UseAllAvailableWidth()
                            .SetMarginBottom(10);
                        
                        tIng.AddHeaderCell(CreateHeaderCell("Concepto/Venta"));
                        tIng.AddHeaderCell(CreateHeaderCell("Monto"));

                        if (day.Ingresos.Any())
                        {
                            foreach (var m in day.Ingresos)
                            {
                                tIng.AddCell(CreateDataCell(m.Concepto));
                                tIng.AddCell(CreateMoneyCell($"S/ {m.Monto:N2}"));
                            }
                        }
                        else
                        {
                            tIng.AddCell(new iText.Layout.Element.Cell(1, 2)
                                .Add(new iText.Layout.Element.Paragraph("Sin ingresos registrados")
                                    .SetFontSize(9)
                                    .SetFontColor(new iText.Kernel.Colors.DeviceRgb(150, 150, 150)))
                                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                                .SetPadding(8));
                        }

                        // Fila de total
                        tIng.AddCell(CreateTotalCell("TOTAL INGRESOS", successBg));
                        tIng.AddCell(CreateTotalMoneyCell($"S/ {day.TotalIngresos:N2}", successBg));
                        
                        dayCard.Add(tIng);

                        // === TABLA DE EGRESOS ===
                        var egrLabel = new iText.Layout.Element.Paragraph("💸 EGRESOS / GASTOS (-)")
                            .SetFontSize(10)
                            .SetFontColor(new iText.Kernel.Colors.DeviceRgb(198, 40, 40))
                            .SetMarginBottom(5);
                        dayCard.Add(egrLabel);

                        var tEgr = new iText.Layout.Element.Table(iText.Layout.Properties.UnitValue.CreatePercentArray(new float[] { 75, 25 }))
                            .UseAllAvailableWidth()
                            .SetMarginBottom(10);
                        
                        tEgr.AddHeaderCell(CreateHeaderCell("Concepto/Responsable"));
                        tEgr.AddHeaderCell(CreateHeaderCell("Monto"));

                        if (day.Egresos.Any())
                        {
                            foreach (var m in day.Egresos)
                            {
                                tEgr.AddCell(CreateDataCell(m.Concepto));
                                tEgr.AddCell(CreateMoneyCell($"S/ {m.Monto:N2}"));
                            }
                        }
                        else
                        {
                            tEgr.AddCell(new iText.Layout.Element.Cell(1, 2)
                                .Add(new iText.Layout.Element.Paragraph("Sin egresos registrados")
                                    .SetFontSize(9)
                                    .SetFontColor(new iText.Kernel.Colors.DeviceRgb(150, 150, 150)))
                                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                                .SetPadding(8));
                        }

                        // Fila de total
                        tEgr.AddCell(CreateTotalCell("TOTAL EGRESOS", errorBg));
                        tEgr.AddCell(CreateTotalMoneyCell($"S/ {day.TotalEgresos:N2}", errorBg));
                        
                        dayCard.Add(tEgr);

                        var balanceCard = new iText.Layout.Element.Div()
                            .SetBackgroundColor(new iText.Kernel.Colors.DeviceRgb(224, 247, 250))
                            .SetBorder(new iText.Layout.Borders.SolidBorder(primaryColor, 1.5f))
                            .SetPadding(10)
                            .SetMarginTop(5);

                        var balanceTitle = new iText.Layout.Element.Paragraph("📊 BALANCE DE CIERRE")
                            .SetFontSize(10)
                            .SetFontColor(primaryColor)
                            .SetMarginBottom(5);
                        balanceCard.Add(balanceTitle);

                        var balanceGrid = new iText.Layout.Element.Table(iText.Layout.Properties.UnitValue.CreatePercentArray(new float[] { 60, 40 }))
                            .UseAllAvailableWidth();
                        
                        balanceGrid.AddCell(CreateBalanceRow("Saldo Inicial (Día anterior)", $"S/ {day.SaldoAyer:N2}"));
                        balanceGrid.AddCell(CreateBalanceRow("Mov. del día (Ing - Egr)", $"S/ {day.TotalDia:N2}", day.TotalDia < 0));
                        balanceGrid.AddCell(CreateBalanceRow("SALDO FINAL EN CAJA", $"S/ {day.SaldoHoy:N2}", false, true));

                        balanceCard.Add(balanceGrid);
                        dayCard.Add(balanceCard);

                        document.Add(dayCard);

                        // Actualizar posición Y aproximada
                        currentY -= estimatedHeightForDay;
                    }

                    // CRÍTICO: Cerrar el document ANTES de agregar footers
                    document.Close();

                    // ===== PIE DE PÁGINA CON NUMERACIÓN =====
                    AddPageNumbers(pdf, primaryColor);

                    // El pdf se cerrará automáticamente por el using statement
                    
                    System.Windows.MessageBox.Show("✅ PDF generado exitosamente.", "Exportación Completa", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"Error al generar PDF: {ex.Message}\n\nDetalle: {ex.StackTrace}", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                }
            }
        }

        // MÉTODOS AUXILIARES PARA PDF
        private iText.Layout.Element.Cell CreateSummaryCell(string label, string value, iText.Kernel.Colors.Color bgcolor)
        {
            var cell = new iText.Layout.Element.Cell();
            cell.Add(new iText.Layout.Element.Paragraph(label).SetFontSize(8).SetMarginBottom(3));
            cell.Add(new iText.Layout.Element.Paragraph(value).SetFontSize(11));
            cell.SetBackgroundColor(bgcolor);
            cell.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);
            cell.SetPadding(8);
            return cell;
        }

        private iText.Layout.Element.Cell CreateHeaderCell(string text)
        {
            return new iText.Layout.Element.Cell()
                .Add(new iText.Layout.Element.Paragraph(text)
                    .SetFontSize(9)
                    .SetFontColor(iText.Kernel.Colors.DeviceRgb.WHITE))
                .SetBackgroundColor(new iText.Kernel.Colors.DeviceRgb(0, 121, 107))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                .SetPadding(6);
        }

        private iText.Layout.Element.Cell CreateDataCell(string text)
        {
            return new iText.Layout.Element.Cell()
                .Add(new iText.Layout.Element.Paragraph(text).SetFontSize(9))
                .SetPadding(5)
                .SetBorder(new iText.Layout.Borders.SolidBorder(new iText.Kernel.Colors.DeviceRgb(220, 220, 220), 0.5f));
        }

        private iText.Layout.Element.Cell CreateMoneyCell(string text)
        {
            return new iText.Layout.Element.Cell()
                .Add(new iText.Layout.Element.Paragraph(text).SetFontSize(9))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                .SetPadding(5)
                .SetBorder(new iText.Layout.Borders.SolidBorder(new iText.Kernel.Colors.DeviceRgb(220, 220, 220), 0.5f));
        }

        private iText.Layout.Element.Cell CreateTotalCell(string text, iText.Kernel.Colors.Color bgcolor)
        {
            return new iText.Layout.Element.Cell()
                .Add(new iText.Layout.Element.Paragraph(text).SetFontSize(9))
                .SetBackgroundColor(bgcolor)
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                .SetPadding(6);
        }

        private iText.Layout.Element.Cell CreateTotalMoneyCell(string text, iText.Kernel.Colors.Color bgcolor)
        {
            return new iText.Layout.Element.Cell()
                .Add(new iText.Layout.Element.Paragraph(text).SetFontSize(10))
                .SetBackgroundColor(bgcolor)
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                .SetPadding(6);
        }

        private iText.Layout.Element.Cell CreateBalanceRow(string label, string value, bool isNegative = false, bool isFinal = false)
        {
            var labelCell = new iText.Layout.Element.Cell()
                .Add(new iText.Layout.Element.Paragraph(label)
                    .SetFontSize(isFinal ? 10 : 9))
                .SetBorder(null)
                .SetPadding(3);

            var valueCell = new iText.Layout.Element.Cell()
                .Add(new iText.Layout.Element.Paragraph(value)
                    .SetFontSize(isFinal ? 11 : 9)
                    .SetFontColor(isNegative ? new iText.Kernel.Colors.DeviceRgb(198, 40, 40) : iText.Kernel.Colors.DeviceRgb.BLACK))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                .SetBorder(null)
                .SetPadding(3);

            if (isFinal)
            {
                var bgcolor = new iText.Kernel.Colors.DeviceRgb(178, 223, 219);
                labelCell.SetBackgroundColor(bgcolor);
                valueCell.SetBackgroundColor(bgcolor);
            }

            return labelCell;
        }

        private void AddPageNumbers(iText.Kernel.Pdf.PdfDocument pdf, iText.Kernel.Colors.Color color)
        {
            int totalPages = pdf.GetNumberOfPages();
            for (int i = 1; i <= totalPages; i++)
            {
                var page = pdf.GetPage(i);
                var pageSize = page.GetPageSize();
                var canvas = new iText.Kernel.Pdf.Canvas.PdfCanvas(page);
                
                // Footer text
                string footerText = $"Página {i} de {totalPages} | Generado: {DateTime.Now:dd/MM/yyyy HH:mm} | Hogar Protegido - Sistema de Tesorería";
                
                canvas.BeginText()
                    .SetFontAndSize(iText.Kernel.Font.PdfFontFactory.CreateFont(iText.IO.Font.Constants.StandardFonts.HELVETICA), 8)
                    .SetColor(color, true)
                    .MoveText(pageSize.GetWidth() / 2 - 150, 20)
                    .ShowText(footerText)
                    .EndText();
            }
        }

        private void ExportarAExcel()
        {
            if (DailyReports.Count == 0) return;

            var sfd = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Archivos de Excel (*.xlsx)|*.xlsx",
                FileName = $"Reporte_Hogar_{DateTime.Now:yyyyMMdd}.xlsx"
            };

            if (sfd.ShowDialog() == true)
            {
                using var workbook = new ClosedXML.Excel.XLWorkbook();
                var ws = workbook.Worksheets.Add("Balance Diario");

                // Headers
                ws.Cell(1, 1).Value = "Fecha";
                ws.Cell(1, 2).Value = "Ingresos";
                ws.Cell(1, 3).Value = "Gastos";
                ws.Cell(1, 4).Value = "Saldo Diario";
                ws.Cell(1, 5).Value = "Caja Total";

                var headerRange = ws.Range(1, 1, 1, 5);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Teal;
                headerRange.Style.Font.FontColor = ClosedXML.Excel.XLColor.White;

                int row = 2;
                foreach (var r in DailyReports)
                {
                    ws.Cell(row, 1).Value = r.Fecha;
                    ws.Cell(row, 2).Value = r.TotalIngresos;
                    ws.Cell(row, 3).Value = r.TotalEgresos;
                    ws.Cell(row, 4).Value = r.TotalDia;
                    ws.Cell(row, 5).Value = r.SaldoHoy;
                    
                    if (r.TotalDia < 0) ws.Cell(row, 4).Style.Font.FontColor = ClosedXML.Excel.XLColor.Red;
                    row++;
                }

                ws.Columns().AdjustToContents();
                ws.Column(1).Width = 15; // Asegurar ancho para la fecha
                ws.Range(2, 2, row - 1, 5).Style.NumberFormat.Format = "[$S/ ]#,##0.00";
                
                workbook.SaveAs(sfd.FileName);
                System.Windows.MessageBox.Show("Excel generado con éxito.");
            }
        }

        public void GenerarReporte()
        {
            DailyReports.Clear();
            
            // 1. OBTENER MOVIMIENTOS (Con o sin filtro)
            var query = _mainViewModel.Movimientos.AsEnumerable();
            if (FilterByDateEnabled)
            {
                if (FechaInicio.HasValue) query = query.Where(m => m.Fecha.Date >= FechaInicio.Value.Date);
                if (FechaFin.HasValue) query = query.Where(m => m.Fecha.Date <= FechaFin.Value.Date);
            }

            var allMovesList = query.ToList();

            // 2. GENERAR REPORTES DIARIOS
            var sortedDays = allMovesList
                .GroupBy(m => m.Fecha.Date)
                .OrderBy(g => g.Key)
                .ToList();

            decimal lastBalance = 0;
            var reportList = new System.Collections.Generic.List<DailyReportItem>();

            foreach (var dayGroup in sortedDays)
            {
                var item = new DailyReportItem
                {
                    Fecha = dayGroup.Key,
                    SaldoAyer = lastBalance
                };

                foreach (var move in dayGroup)
                {
                    if (move.Tipo == TipoMovimiento.Ingreso) item.Ingresos.Add(move);
                    else item.Egresos.Add(move);
                }

                reportList.Add(item);
                lastBalance = item.SaldoHoy;
            }

            reportList.Reverse();
            foreach(var item in reportList) DailyReports.Add(item);

        }
    }
}
