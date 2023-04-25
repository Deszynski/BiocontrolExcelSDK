using ClosedXML.Excel;
using Microsoft.Win32;
using Soneta.Business;
using Soneta.Handel;
using System;

[assembly: Worker(typeof(BiocontrolExcelSDK.EksportZamowienia), typeof(DokHandlowe))]

namespace BiocontrolExcelSDK
{
    internal class EksportZamowienia
    {
        [Context]
        public Context Context { get; set; }

        [Action(
            "Open Sales Orders",
            Priority = 60,
            Icon = ActionIcon.Copy,
            Mode = ActionMode.SingleSession,
            Target = ActionTarget.Menu | ActionTarget.ToolbarWithText)]

        public void MyActionZamowienia()
        {
            HandelModule handelModule;
            View zamowienia;

            string fileName = "Open Sales Orders " + DateTime.Now.ToString().Remove(10) + ".xlsx";
            double[] colWidth = new double[] { 15, 15, 13, 13, 35, 7, 10, 50, 8, 8, 8, 11, 13, 13, 12, 9, 9, 10 };

            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                FilterIndex = 2,
                RestoreDirectory = true,
                InitialDirectory = @"C:\Users\" + Environment.UserName + @"\OneDrive\Dokumenty\",
                FileName = fileName
            };

            bool? result = saveDialog.ShowDialog();

            if (result == true)
            {
                string path = saveDialog.FileName;

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Open Sales Orders");

                    #region headlines
                    worksheet.Cell(1, 1).Value = "Typ dokumentu";
                    worksheet.Cell(1, 2).Value = "Nr dokumentu";
                    worksheet.Cell(1, 3).Value = "Stan";
                    worksheet.Cell(1, 4).Value = "Nr nabywcy (sprzedaż)";
                    worksheet.Cell(1, 5).Value = "Nazwa nabywcy";
                    worksheet.Cell(1, 6).Value = "Typ";
                    worksheet.Cell(1, 7).Value = "Nr";
                    worksheet.Cell(1, 8).Value = "Pełna nazwa";
                    worksheet.Cell(1, 9).Value = "Kod lokalizacji";
                    worksheet.Cell(1, 10).Value = "Ilość";
                    worksheet.Cell(1, 11).Value = "Kod jednostki miary";
                    worksheet.Cell(1, 12).Value = "Ilość zarezerw. (podst. JM)";
                    worksheet.Cell(1, 13).Value = "Cena jedn. z rabatem Bez VAT";
                    worksheet.Cell(1, 14).Value = "Kwota wiersza Bez VAT";
                    worksheet.Cell(1, 15).Value = "Data wydania";
                    worksheet.Cell(1, 16).Value = "Ilość pozostała";
                    worksheet.Cell(1, 17).Value = "Ilość do wydania";
                    worksheet.Cell(1, 18).Value = "Nr kampanii";

                    worksheet.Range(1, 1, 1, 18).Style.Fill.BackgroundColor = XLColor.LightGray;
                    #endregion

                    // eksport zamowien sprzedazy
                    int currentRow = 2;

                    using (Session session = Context.Login.CreateSession(false, false))
                    {
                        handelModule = HandelModule.GetInstance(session);
                        zamowienia = handelModule.DokHandlowe.CreateView();
                        zamowienia.Condition &= new FieldCondition.Equal("Definicja", handelModule.DefDokHandlowych.ZamówienieOdbiorcy);

                        foreach (DokumentHandlowy zo in zamowienia)
                        {
                            foreach (PozycjaDokHandlowego poz in zo.Pozycje) if (poz.IlośćZasobu.Value != 0)
                            {
                                worksheet.Cell(currentRow, 1).Value = "Zamówienie";
                                worksheet.Cell(currentRow, 2).Value = zo.Numer.ToString();
                                worksheet.Cell(currentRow, 3).Value = zo.Stan.ToString();
                                worksheet.Cell(currentRow, 4).Value = zo.Kontrahent.Kod.ToString();
                                worksheet.Cell(currentRow, 5).Value = zo.Kontrahent.Nazwa.ToString();
                                worksheet.Cell(currentRow, 6).Value = "";
                                worksheet.Cell(currentRow, 7).Value = poz.Towar.Kod.ToString();
                                worksheet.Cell(currentRow, 8).Value = poz.Towar.Nazwa.ToString();
                                worksheet.Cell(currentRow, 9).Value = zo.Magazyn.ToString();
                                worksheet.Cell(currentRow, 10).Value = poz.Ilosc.Value.ToString();
                                worksheet.Cell(currentRow, 11).Value = poz.Towar.Jednostka.ToString();
                                worksheet.Cell(currentRow, 12).Value = "0";
                                worksheet.Cell(currentRow, 13).Value = poz.CenaJednostkowa.Value;
                                worksheet.Cell(currentRow, 14).Value = poz.CenaJednostkowa.Value * poz.Ilosc.Value;
                                worksheet.Cell(currentRow, 15).Value = poz.Features["Data realizacji"].ToString();
                                worksheet.Cell(currentRow, 16).Value = (int)poz.IlośćZasobu.Value;
                                worksheet.Cell(currentRow, 17).Value = (int)poz.IlośćZasobu.Value;
                                worksheet.Cell(currentRow, 18).Value = "";

                                currentRow++;
                            }
                        }

                        session.Save();
                    }

                    #region total
                    worksheet.Cell(currentRow, 14).FormulaA1 = "=SUM(N2:N" + (currentRow - 1) + ")";
                    worksheet.Cell(currentRow, 14).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    worksheet.Cell(currentRow, 14).Style.Fill.BackgroundColor = XLColor.LightGray;
                    #endregion

                    #region worksheet style
                    for (int row = 1; row < currentRow; row++)
                    {
                        worksheet.Row(row).Height = 13;
                        for (int col = 1; col <= 18; col++)
                            worksheet.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    }

                    worksheet.Column(4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Column(9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Column(14).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Column(16).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // column number format
                    worksheet.Column(13).Style.NumberFormat.Format = "0.00";
                    worksheet.Column(14).Style.NumberFormat.Format = "0.00";

                    worksheet.Rows().Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet.Row(1).Height = 23;
                    worksheet.Row(1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                    worksheet.Range(1, 1, 1, 17).SetAutoFilter();
                    worksheet.SheetView.FreezeRows(1);
                    worksheet.Columns().Style.Font.SetFontName("Arial");
                    worksheet.Columns().Style.Font.SetFontSize(10);
                    for (int i = 1; i <= 18; i++)
                        worksheet.Column(i).Width = colWidth[i - 1];

                    worksheet.Column(6).Hide();
                    worksheet.Column(18).Hide();
                    #endregion

                    // zapis do pliku
                    workbook.SaveAs(path);
                }
            }                
        }
    }
}
