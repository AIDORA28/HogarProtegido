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
                    using var document = new iText.Layout.Document(pdf, iText.Kernel.Geom.PageSize.A4);
                    
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
                            .SetBorder(new iText.Layout.Borders.SolidBorder(accentColor, 2f))
                            .SetPadding(18)
                            .SetMarginBottom(20)
                            .SetKeepTogether(true);

                        // Encabezado de la fecha
                        var dateHeader = new iText.Layout.Element.Paragraph($"📅 {day.Fecha:dddd, dd 'de' MMMM 'de' yyyy}".ToUpper())
                            .SetFontSize(14)
                            .SetFontColor(primaryColor)
                            .SetMarginBottom(12);
                        dayCard.Add(dateHeader);

                        // === TABLA DE INGRESOS ===
                        var ingLabel = new iText.Layout.Element.Paragraph("💰 INGRESOS (+)")
                            .SetFontSize(11)
                            .SetFontColor(new iText.Kernel.Colors.DeviceRgb(46, 125, 50))
                            .SetMarginBottom(6);
                        dayCard.Add(ingLabel);

                        var tIng = new iText.Layout.Element.Table(iText.Layout.Properties.UnitValue.CreatePercentArray(new float[] { 75, 25 }))
                            .UseAllAvailableWidth()
                            .SetMarginBottom(12);
                        
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
                            .SetFontSize(11)
                            .SetFontColor(new iText.Kernel.Colors.DeviceRgb(198, 40, 40))
                            .SetMarginBottom(6);
                        dayCard.Add(egrLabel);

                        var tEgr = new iText.Layout.Element.Table(iText.Layout.Properties.UnitValue.CreatePercentArray(new float[] { 75, 25 }))
                            .UseAllAvailableWidth()
                            .SetMarginBottom(12);
                        
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
                            .SetBorder(new iText.Layout.Borders.SolidBorder(primaryColor, 2f))
                            .SetPadding(12)
                            .SetMarginTop(8);

                        var balanceTitle = new iText.Layout.Element.Paragraph("📊 BALANCE DE CIERRE")
                            .SetFontSize(11)
                            .SetFontColor(primaryColor)
                            .SetMarginBottom(6);
                        balanceCard.Add(balanceTitle);

                        var balanceGrid = new iText.Layout.Element.Table(iText.Layout.Properties.UnitValue.CreatePercentArray(new float[] { 60, 40 }))
                            .UseAllAvailableWidth();
                        
                        CreateBalanceRow(balanceGrid, "Saldo Inicial (Día anterior)", $"S/ {day.SaldoAyer:N2}");
                        CreateBalanceRow(balanceGrid, "Mov. del día (Ing - Egr)", $"S/ {day.TotalDia:N2}", day.TotalDia < 0);
                        CreateBalanceRow(balanceGrid, "SALDO FINAL EN CAJA", $"S/ {day.SaldoHoy:N2}", false, true);

                        balanceCard.Add(balanceGrid);
                        dayCard.Add(balanceCard);

                        document.Add(dayCard);

                        // Actualizar posición Y aproximada
                        currentY -= estimatedHeightForDay;
                    }


                    // NOTA: No llamar a document.Close() manualmente ni a AddPageNumbers()
                    // porque causa conflictos con iText7. El using statement cierra automáticamente.
                    // TODO: Implementar footers usando PageEvent handlers en futuras versiones.
                    
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

        private void CreateBalanceRow(iText.Layout.Element.Table table, string label, string value, bool isNegative = false, bool isFinal = false)
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

            // Agregar AMBAS celdas a la tabla
            table.AddCell(labelCell);
            table.AddCell(valueCell);
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
                var ws = workbook.Worksheets.Add("Reporte Detallado");

                // Configurar página para impresión
                ws.PageSetup.PageOrientation = ClosedXML.Excel.XLPageOrientation.Portrait;
                ws.PageSetup.PaperSize = ClosedXML.Excel.XLPaperSize.A4Paper;
                ws.PageSetup.Margins.Left = 0.5;
                ws.PageSetup.Margins.Right = 0.5;

                int currentRow = 1;

                // ===== TÍTULO PRINCIPAL =====
                ws.Cell(currentRow, 1).Value = "HOGAR PROTEGIDO - REPORTE DETALLADO DE CAJA";
                ws.Range(currentRow, 1, currentRow, 5).Merge();
                ws.Cell(currentRow, 1).Style.Font.FontSize = 16;
                ws.Cell(currentRow, 1).Style.Font.Bold = true;
                ws.Cell(currentRow, 1).Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
                ws.Cell(currentRow, 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(0, 121, 107);
                ws.Cell(currentRow, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                ws.Cell(currentRow, 1).Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;
                ws.Row(currentRow).Height = 30;
                currentRow += 2;

                // ===== RESUMEN EJECUTIVO =====
                var totalIngresos = DailyReports.Sum(d => d.TotalIngresos);
                var totalEgresos = DailyReports.Sum(d => d.TotalEgresos);
                var saldoFinal = DailyReports.FirstOrDefault()?.SaldoHoy ?? 0; // Primer día (más reciente)

                ws.Cell(currentRow, 1).Value = "Días Reportados:";
                ws.Cell(currentRow, 1).Style.Font.Bold = true;
                ws.Cell(currentRow, 2).Value = DailyReports.Count;

                ws.Cell(currentRow, 3).Value = "Total Ingresos:";
                ws.Cell(currentRow, 3).Style.Font.Bold = true;
                ws.Cell(currentRow, 4).Value = totalIngresos;
                ws.Cell(currentRow, 4).Style.NumberFormat.Format = "[$S/ ]#,##0.00";
                ws.Cell(currentRow, 4).Style.Font.FontColor = ClosedXML.Excel.XLColor.Green;

                currentRow++;
                ws.Cell(currentRow, 3).Value = "Total Egresos:";
                ws.Cell(currentRow, 3).Style.Font.Bold = true;
                ws.Cell(currentRow, 4).Value = totalEgresos;
                ws.Cell(currentRow, 4).Style.NumberFormat.Format = "[$S/ ]#,##0.00";
                ws.Cell(currentRow, 4).Style.Font.FontColor = ClosedXML.Excel.XLColor.Red;

                currentRow++;
                ws.Cell(currentRow, 3).Value = "Saldo Final en Caja:";
                ws.Cell(currentRow, 3).Style.Font.Bold = true;
                ws.Cell(currentRow, 3).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(178, 223, 219);
                ws.Cell(currentRow, 4).Value = saldoFinal;
                ws.Cell(currentRow, 4).Style.NumberFormat.Format = "[$S/ ]#,##0.00";
                ws.Cell(currentRow, 4).Style.Font.Bold = true;
                ws.Cell(currentRow, 4).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(178, 223, 219);

                currentRow += 3;

                // ===== DETALLE POR DÍA =====
                foreach (var day in DailyReports)
                {
                    // ENCABEZADO DEL DÍA
                    ws.Cell(currentRow, 1).Value = day.Fecha.ToString("dddd, dd 'de' MMMM 'de' yyyy").ToUpper();
                    ws.Range(currentRow, 1, currentRow, 5).Merge();
                    ws.Cell(currentRow, 1).Style.Font.Bold = true;
                    ws.Cell(currentRow, 1).Style.Font.FontSize = 12;
                    ws.Cell(currentRow, 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightGray;
                    ws.Cell(currentRow, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Left;
                    ws.Cell(currentRow, 1).Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Medium;
                    ws.Row(currentRow).Height = 25;
                    currentRow++;

                    // TABLA DE INGRESOS
                    if (day.Ingresos.Any())
                    {
                        ws.Cell(currentRow, 1).Value = "INGRESOS (+)";
                        ws.Cell(currentRow, 1).Style.Font.Bold = true;
                        ws.Cell(currentRow, 1).Style.Font.FontColor = ClosedXML.Excel.XLColor.DarkGreen;
                        currentRow++;

                        // Headers de ingresos
                        ws.Cell(currentRow, 1).Value = "Concepto/Venta";
                        ws.Cell(currentRow, 2).Value = "Monto";
                        ws.Range(currentRow, 1, currentRow, 2).Style.Font.Bold = true;
                        ws.Range(currentRow, 1, currentRow, 2).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(220, 237, 200);
                        ws.Range(currentRow, 1, currentRow, 2).Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                        currentRow++;

                        // Datos de ingresos
                        foreach (var mov in day.Ingresos)
                        {
                            ws.Cell(currentRow, 1).Value = mov.Concepto;
                            ws.Cell(currentRow, 2).Value = mov.Monto;
                            ws.Cell(currentRow, 2).Style.NumberFormat.Format = "[$S/ ]#,##0.00";
                            ws.Range(currentRow, 1, currentRow, 2).Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Hair;
                            currentRow++;
                        }

                        // Total ingresos
                        ws.Cell(currentRow, 1).Value = "TOTAL INGRESOS";
                        ws.Cell(currentRow, 1).Style.Font.Bold = true;
                        ws.Cell(currentRow, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Right;
                        ws.Cell(currentRow, 2).Value = day.TotalIngresos;
                        ws.Cell(currentRow, 2).Style.NumberFormat.Format = "[$S/ ]#,##0.00";
                        ws.Cell(currentRow, 2).Style.Font.Bold = true;
                        ws.Range(currentRow, 1, currentRow, 2).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(220, 237, 200);
                        ws.Range(currentRow, 1, currentRow, 2).Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Medium;
                        currentRow++;
                    }
                    else
                    {
                        ws.Cell(currentRow, 1).Value = "INGRESOS (+): Sin ingresos registrados";
                        ws.Cell(currentRow, 1).Style.Font.Italic = true;
                        ws.Cell(currentRow, 1).Style.Font.FontColor = ClosedXML.Excel.XLColor.Gray;
                        currentRow++;
                    }

                    currentRow++; // Espacio entre secciones

                    // TABLA DE EGRESOS
                    if (day.Egresos.Any())
                    {
                        ws.Cell(currentRow, 1).Value = "EGRESOS / GASTOS (-)";
                        ws.Cell(currentRow, 1).Style.Font.Bold = true;
                        ws.Cell(currentRow, 1).Style.Font.FontColor = ClosedXML.Excel.XLColor.DarkRed;
                        currentRow++;

                        // Headers de egresos
                        ws.Cell(currentRow, 1).Value = "Concepto/Responsable";
                        ws.Cell(currentRow, 2).Value = "Monto";
                        ws.Range(currentRow, 1, currentRow, 2).Style.Font.Bold = true;
                        ws.Range(currentRow, 1, currentRow, 2).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(255, 205, 210);
                        ws.Range(currentRow, 1, currentRow, 2).Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                        currentRow++;

                        // Datos de egresos
                        foreach (var mov in day.Egresos)
                        {
                            ws.Cell(currentRow, 1).Value = mov.Concepto;
                            ws.Cell(currentRow, 2).Value = mov.Monto;
                            ws.Cell(currentRow, 2).Style.NumberFormat.Format = "[$S/ ]#,##0.00";
                            ws.Range(currentRow, 1, currentRow, 2).Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Hair;
                            currentRow++;
                        }

                        // Total egresos
                        ws.Cell(currentRow, 1).Value = "TOTAL EGRESOS";
                        ws.Cell(currentRow, 1).Style.Font.Bold = true;
                        ws.Cell(currentRow, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Right;
                        ws.Cell(currentRow, 2).Value = day.TotalEgresos;
                        ws.Cell(currentRow, 2).Style.NumberFormat.Format = "[$S/ ]#,##0.00";
                        ws.Cell(currentRow, 2).Style.Font.Bold = true;
                        ws.Range(currentRow, 1, currentRow, 2).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(255, 205, 210);
                        ws.Range(currentRow, 1, currentRow, 2).Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Medium;
                        currentRow++;
                    }
                    else
                    {
                        ws.Cell(currentRow, 1).Value = "EGRESOS (-): Sin egresos registrados";
                        ws.Cell(currentRow, 1).Style.Font.Italic = true;
                        ws.Cell(currentRow, 1).Style.Font.FontColor = ClosedXML.Excel.XLColor.Gray;
                        currentRow++;
                    }

                    currentRow++; // Espacio entre secciones

                    // BALANCE DE CIERRE
                    ws.Cell(currentRow, 1).Value = "BALANCE DE CIERRE";
                    ws.Range(currentRow, 1, currentRow, 2).Merge();
                    ws.Cell(currentRow, 1).Style.Font.Bold = true;
                    ws.Cell(currentRow, 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(224, 247, 250);
                    ws.Cell(currentRow, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                    ws.Cell(currentRow, 1).Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Medium;
                    currentRow++;

                    ws.Cell(currentRow, 1).Value = "Saldo Inicial (Día anterior)";
                    ws.Cell(currentRow, 2).Value = day.SaldoAyer;
                    ws.Cell(currentRow, 2).Style.NumberFormat.Format = "[$S/ ]#,##0.00";
                    currentRow++;

                    ws.Cell(currentRow, 1).Value = "Movimiento del día (Ing - Egr)";
                    ws.Cell(currentRow, 2).Value = day.TotalDia;
                    ws.Cell(currentRow, 2).Style.NumberFormat.Format = "[$S/ ]#,##0.00";
                    if (day.TotalDia < 0) ws.Cell(currentRow, 2).Style.Font.FontColor = ClosedXML.Excel.XLColor.Red;
                    currentRow++;

                    ws.Cell(currentRow, 1).Value = "SALDO FINAL EN CAJA";
                    ws.Cell(currentRow, 1).Style.Font.Bold = true;
                    ws.Cell(currentRow, 2).Value = day.SaldoHoy;
                    ws.Cell(currentRow, 2).Style.NumberFormat.Format = "[$S/ ]#,##0.00";
                    ws.Cell(currentRow, 2).Style.Font.Bold = true;
                    ws.Range(currentRow, 1, currentRow, 2).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(178, 223, 219);
                    ws.Range(currentRow, 1, currentRow, 2).Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Medium;
                    currentRow++;

                    currentRow += 2; // Espacio grande entre días
                }

                // Ajustar anchos de columna
                ws.Column(1).Width = 40; // Concepto
                ws.Column(2).Width = 15; // Monto
                ws.Column(3).Width = 20; // Etiquetas resumen
                ws.Column(4).Width = 15; // Valores resumen

                try
                {
                    workbook.SaveAs(sfd.FileName);
                    System.Windows.MessageBox.Show("✅ Excel generado exitosamente con formato profesional.", "Exportación Completa", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                }
                catch (System.IO.IOException ex) when (ex.HResult == unchecked((int)0x80070020))
                {
                    System.Windows.MessageBox.Show(
                        $"⚠️ No se puede guardar el archivo porque ya está abierto.\n\n" +
                        $"Por favor:\n" +
                        $"1. Cierra el archivo '{System.IO.Path.GetFileName(sfd.FileName)}' en Excel\n" +
                        $"2. O elige un nombre diferente para el archivo\n\n" +
                        $"Luego intenta exportar nuevamente.",
                        "Archivo en Uso",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Warning);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show(
                        $"Error al generar Excel: {ex.Message}",
                        "Error",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Error);
                }
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

            // Invertir para mostrar días más recientes primero
            reportList.Reverse();
            
            // RECALCULAR SaldoAyer después del reverse para mantener continuidad
            for (int i = 0; i < reportList.Count; i++)
            {
                if (i == reportList.Count - 1)
                {
                    reportList[i].SaldoAyer = 0; // El día más antiguo (ahora al final) empieza en 0
                }
                else
                {
                    reportList[i].SaldoAyer = reportList[i + 1].SaldoHoy; // Saldo inicial = saldo final del día siguiente (que cronológicamente es el anterior)
                }
            }
            
            foreach(var item in reportList) DailyReports.Add(item);

        }
    }
}
