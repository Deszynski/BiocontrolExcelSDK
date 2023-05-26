using Microsoft.Win32;
using Soneta.Business;
using Soneta.Handel;
using Soneta.Ksiega;
using System.Collections.Generic;
using System;
using Soneta.Types;
using ClosedXML.Excel;
using System.Globalization;

[assembly: Worker(typeof(BiocontrolExcelSDK.EksportFaktury), typeof(DokHandlowe))]

namespace BiocontrolExcelSDK
{
    internal class EksportFaktury
    {
        [Context]
        public Context Context { get; set; }

        [Action(
            "Sales invoices",
            Priority = 30,
            Icon = ActionIcon.Copy,
            Mode = ActionMode.Progress,
            Target = ActionTarget.Menu | ActionTarget.ToolbarWithText)]

        public void MyActionFaktury()
        {
            HandelModule handelModule;
            List<DokumentHandlowy> fakturyList = new List<DokumentHandlowy>();

            string fileName = "Sales invoices " + DateTime.Now.ToString().Remove(10) + ".xlsx";

            #region zmienne
            View faktury;
            View elementy;

            // zmienne arkusz 1
            int currentRow = 3;
            int firstRowOfMonth = 3;
            int endRow = 0;
            Date dataOld = new Date();
            Date dataNew = new Date();
            double monthTotalEUR = 0, monthTotalNOK = 0, monthTotalPLN = 0;
            double totalEUR = 0, totalNOK = 0, totalPLN = 0;
            string monthStr = "";

            // zmienne arkusz 2
            int currentRow2 = 3;

            // zmienne arkusz 3
            int currentRow3 = 2;
            List<string> dim3 = new List<string>();
            KsiegaModule ksiegaModule;
            ElemSlownika e;

            // zmienne arkusz 4
            int currentRow4 = 2;
            List<string> dim4 = new List<string>();
            #endregion

            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                FilterIndex = 2,
                RestoreDirectory = true,
                InitialDirectory = @"C:\Users\" + Environment.UserName + @"\OneDrive\Dokumenty\",
                FileName = fileName
            };

            List<string> miesiace = new List<string>();

            bool? result = saveDialog.ShowDialog();

            if (result == true)
            {
                string path = saveDialog.FileName;

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("HEADER");
                    var worksheet2 = workbook.Worksheets.Add("DETAILS");
                    var worksheet3 = workbook.Worksheets.Add("DIMENSION");
                    var worksheet4 = workbook.Worksheets.Add("DIMENSION WITHOUT BCN");

                    #region arkusz 1,2
                    using (Session session = Context.Login.CreateSession(false, false))
                    {
                        handelModule = HandelModule.GetInstance(session);

                        INavigatorContext inc = Context[typeof(INavigatorContext)] as INavigatorContext;

                        if (inc != null && inc.SelectedRows.Length != 0)
                        {
                            faktury = handelModule.DokHandlowe.CreateView();
                            faktury.Condition &= new FieldCondition.In("Definicja", handelModule.DefDokHandlowych.FakturaSprzedaży, handelModule.DefDokHandlowych.KorektaSprzedaży);
                            faktury.Sort = "Data";

                            foreach (object o in inc.SelectedRows)
                            {
                                DokumentHandlowy faktura = (DokumentHandlowy)o;
                                fakturyList.Add(faktura);
                            }
                        }
                        else
                        {
                            faktury = handelModule.DokHandlowe.CreateView();
                            faktury.Condition &= new FieldCondition.In("Definicja", handelModule.DefDokHandlowych.FakturaSprzedaży, handelModule.DefDokHandlowych.KorektaSprzedaży);
                            faktury.Sort = "Data";
                        }

                        foreach (DokumentHandlowy fv in fakturyList.Count == 0 ? faktury.ToArray<DokumentHandlowy>() : fakturyList.ToArray())
                        {
                            #region arkusz 1
                            dataOld = dataNew;
                            dataNew = fv.Data;
                            string temp = (CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(dataNew.Month) + " " + dataNew.Year).ToUpper();
                            if (!miesiace.Contains(temp))
                                miesiace.Add(temp);

                            worksheet.Cell(currentRow, 1).Value = fv.Numer.ToString();
                            worksheet.Cell(currentRow, 2).Value = fv.Data.ToString();
                            worksheet.Cell(currentRow, 3).Value = fv.Kontrahent.Kod.ToString();
                            worksheet.Cell(currentRow, 4).Value = fv.Kontrahent.Nazwa.ToString();
                            worksheet.Cell(currentRow, 5).Value = fv.WalutaKontrahenta.ToString();
                            worksheet.Cell(currentRow, 6).Value = (double)fv.Suma.NettoCy.Value;
                            worksheet.Cell(currentRow, 7).Value = fv.KursWaluty;
                            worksheet.Cell(currentRow, 8).Value = FakturaOplacona(fv) ? "Tak" : "Nie";
                            if (fv.Platnosci.GetFirst() != null)
                                worksheet.Cell(currentRow, 9).Value = fv.Platnosci.GetFirst().Termin.ToString();

                            if (fv.WalutaKontrahenta.ToString() == "PLN")
                                worksheet.Range(currentRow, 1, currentRow, 9).Style.Fill.BackgroundColor = XLColor.Yellow;
                            else if (fv.WalutaKontrahenta.ToString() == "NOK")
                                worksheet.Range(currentRow, 1, currentRow, 9).Style.Fill.BackgroundColor = XLColor.LightBlue;

                            if (dataOld.Month != dataNew.Month && currentRow >= 7)
                            {
                                #region month total
                                for (int row = firstRowOfMonth; row <= currentRow - 1; row++)
                                {
                                    if (worksheet.Cell(row, 5).Value.ToString() == "EUR")
                                        monthTotalEUR += (double)worksheet.Cell(row, 6).Value;
                                    else if (worksheet.Cell(row, 5).Value.ToString() == "NOK")
                                        monthTotalNOK += (double)worksheet.Cell(row, 6).Value;
                                    else if (worksheet.Cell(row, 5).Value.ToString() == "PLN")
                                        monthTotalPLN += (double)worksheet.Cell(row, 6).Value;
                                }

                                // row 1
                                monthStr = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(dataOld.Month);
                                worksheet.Cell(currentRow - 4, 12).Value = monthStr.ToUpper() + " " + dataOld.Year;
                                worksheet.Cell(currentRow - 4, 13).Value = monthTotalEUR;
                                worksheet.Cell(currentRow - 4, 14).Value = "EUR";
                                worksheet.Range(currentRow - 4, 12, currentRow - 4, 14).Style.Fill.BackgroundColor = XLColor.Yellow;

                                // row 2
                                worksheet.Cell(currentRow - 3, 13).Value = monthTotalNOK;
                                worksheet.Cell(currentRow - 3, 14).Value = "NOK";
                                worksheet.Range(currentRow - 3, 13, currentRow - 3, 14).Style.Fill.BackgroundColor = XLColor.LightBlue;

                                // row 3
                                worksheet.Cell(currentRow - 2, 13).Value = monthTotalPLN;
                                worksheet.Cell(currentRow - 2, 14).Value = "PLN";
                                worksheet.Range(currentRow - 2, 13, currentRow - 2, 14).Style.Fill.BackgroundColor = XLColor.LightYellow;

                                // row 4
                                worksheet.Cell(currentRow - 1, 14).Value = "EUR";
                                worksheet.Range(currentRow - 1, 13, currentRow - 1, 14).Style.Fill.BackgroundColor = XLColor.LightGray;

                                // borders
                                for (int row = currentRow - 4; row <= currentRow - 1; row++)
                                    for (int col = 13; col <= 14; col++)
                                        worksheet.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                                worksheet.Cell(currentRow - 4, 12).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

                                firstRowOfMonth = currentRow;

                                totalEUR += monthTotalEUR;
                                totalNOK += monthTotalNOK;
                                totalPLN += monthTotalPLN;
                                monthTotalEUR = 0;
                                monthTotalNOK = 0;
                                monthTotalPLN = 0;
                                #endregion
                            }

                            currentRow++;
                            #endregion

                            #region arkusz 2
                            foreach (PozycjaDokHandlowego poz in fv.Pozycje)
                            {
                                ElemSlownika elem = (ElemSlownika)poz.Towar.Features["SALES"];
                                worksheet2.Cell(currentRow2, 1).Value = fv.Numer.ToString();
                                worksheet2.Cell(currentRow2, 2).Value = fv.Data.ToString();
                                worksheet2.Cell(currentRow2, 3).Value = fv.Kontrahent.Nazwa.ToString();
                                worksheet2.Cell(currentRow2, 4).Value = poz.Towar.Kod.ToString();
                                worksheet2.Cell(currentRow2, 5).Value = poz.Towar.Nazwa.ToString();
                                worksheet2.Cell(currentRow2, 6).Value = poz.Towar.Jednostka.ToString();
                                worksheet2.Cell(currentRow2, 7).Value = poz.Ilosc.Value.ToString();
                                worksheet2.Cell(currentRow2, 8).Value = poz.CenaPoRabacie.Value * poz.Ilosc.Value;
                                worksheet2.Cell(currentRow2, 9).Value = poz.CenaPoRabacie.Value;
                                worksheet2.Cell(currentRow2, 10).Value = fv.WalutaKontrahenta.ToString();

                                if (elem != null)
                                {
                                    worksheet2.Cell(currentRow2, 11).Value = elem.Symbol.ToString();
                                    worksheet2.Cell(currentRow2, 12).Value = elem.Nazwa.ToString();
                                }

                                worksheet2.Cell(currentRow2, 12 + fv.Data.Month).Value = poz.CenaPoRabacie.Value * poz.Ilosc.Value;

                                if (fv.WalutaKontrahenta.ToString() == "PLN")
                                    worksheet2.Range(currentRow2, 1, currentRow2, 24).Style.Fill.BackgroundColor = XLColor.Yellow;
                                else if (fv.WalutaKontrahenta.ToString() == "NOK")
                                    worksheet2.Range(currentRow2, 1, currentRow2, 24).Style.Fill.BackgroundColor = XLColor.LightBlue;

                                currentRow2++;
                            }
                            #endregion
                        }

                        session.Save();
                    }

                    endRow = currentRow;

                    #region month total (last month)               
                    for (int row = firstRowOfMonth; row <= currentRow - 1; row++)
                    {
                        if (worksheet.Cell(row, 5).Value.ToString() == "EUR")
                            monthTotalEUR += (double)worksheet.Cell(row, 6).Value;
                        else if (worksheet.Cell(row, 5).Value.ToString() == "NOK")
                            monthTotalNOK += (double)worksheet.Cell(row, 6).Value;
                        else if (worksheet.Cell(row, 5).Value.ToString() == "PLN")
                            monthTotalPLN += (double)worksheet.Cell(row, 6).Value;
                    }

                    if (currentRow - firstRowOfMonth < 5)
                        currentRow += 5;

                    // row 1
                    monthStr = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(dataNew.Month);
                    worksheet.Cell(currentRow - 4, 12).Value = monthStr.ToUpper() + " " + dataNew.Year;
                    worksheet.Cell(currentRow - 4, 13).Value = monthTotalEUR;
                    worksheet.Cell(currentRow - 4, 14).Value = "EUR";
                    worksheet.Range(currentRow - 4, 12, currentRow - 4, 14).Style.Fill.BackgroundColor = XLColor.Yellow;

                    // row 2
                    worksheet.Cell(currentRow - 3, 13).Value = monthTotalNOK;
                    worksheet.Cell(currentRow - 3, 14).Value = "NOK";
                    worksheet.Range(currentRow - 3, 13, currentRow - 3, 14).Style.Fill.BackgroundColor = XLColor.LightBlue;

                    // row 3
                    worksheet.Cell(currentRow - 2, 13).Value = monthTotalPLN;
                    worksheet.Cell(currentRow - 2, 14).Value = "PLN";
                    worksheet.Range(currentRow - 2, 13, currentRow - 2, 14).Style.Fill.BackgroundColor = XLColor.LightYellow;

                    // row 4
                    worksheet.Cell(currentRow - 1, 14).Value = "EUR";
                    worksheet.Range(currentRow - 1, 13, currentRow - 1, 14).Style.Fill.BackgroundColor = XLColor.LightGray;

                    // borders
                    for (int row = currentRow - 4; row <= currentRow - 1; row++)
                        for (int col = 13; col <= 14; col++)
                            worksheet.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    worksheet.Cell(currentRow - 4, 12).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

                    firstRowOfMonth = currentRow;

                    totalEUR += monthTotalEUR;
                    totalNOK += monthTotalNOK;
                    totalPLN += monthTotalPLN;
                    #endregion

                    #region total w1
                    worksheet.Cell(endRow, 5).Value = "Sum:";
                    worksheet.Cell(endRow, 6).FormulaA1 = "=SUM(F3:F" + (endRow - 1) + ")";

                    // row 1
                    worksheet.Cell(currentRow + 1, 12).Value = "SUM:";
                    worksheet.Cell(currentRow + 1, 13).Value = totalEUR;
                    worksheet.Cell(currentRow + 1, 14).Value = "EUR";
                    worksheet.Range(currentRow + 1, 12, currentRow + 1, 14).Style.Fill.BackgroundColor = XLColor.Yellow;

                    // row 2
                    worksheet.Cell(currentRow + 2, 13).Value = totalNOK;
                    worksheet.Cell(currentRow + 2, 14).Value = "NOK";
                    worksheet.Range(currentRow + 2, 12, currentRow + 2, 14).Style.Fill.BackgroundColor = XLColor.LightBlue;

                    // row 3
                    worksheet.Cell(currentRow + 3, 13).Value = totalPLN;
                    worksheet.Cell(currentRow + 3, 14).Value = "PLN";
                    worksheet.Range(currentRow + 3, 12, currentRow + 3, 14).Style.Fill.BackgroundColor = XLColor.LightYellow;

                    // row 4
                    worksheet.Cell(currentRow + 4, 14).Value = "EUR";
                    worksheet.Range(currentRow + 4, 12, currentRow + 4, 14).Style.Fill.BackgroundColor = XLColor.LightGray;

                    // borders
                    for (int row = currentRow + 1; row <= currentRow + 4; row++)
                        for (int col = 12; col <= 14; col++)
                            worksheet.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    worksheet.Range(currentRow + 1, 12, currentRow + 4, 14).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
                    #endregion

                    #region total w2
                    for (int col = 1; col <= 12; col++)
                        worksheet2.Cell(currentRow2, col + 12).FormulaA1 = "=SUM(" + worksheet2.Cell(3, col + 12).Address + ":" + worksheet2.Cell(currentRow2 - 1, col + 12).Address + ")";
                    worksheet2.Range(currentRow2, 13, currentRow2, 24).Style.Fill.BackgroundColor = XLColor.Gray;
                    #endregion
                    #endregion

                    #region arkusz 3
                    using (Session s = Context.Login.CreateSession(false, false))
                    {
                        ksiegaModule = KsiegaModule.GetInstance(s);
                        elementy = ksiegaModule.ElemSlownikow.CreateView();
                        elementy.Condition &= new FieldCondition.Equal("Definicja", "Konta Norweskie");

                        foreach (ElemSlownika elem in elementy)
                        {
                            if ((bool)elem.Features["Sal"] == true)
                            {
                                dim3.Add(elem.Symbol.ToString());
                                worksheet3.Cell(currentRow3, 1).Value = elem.Symbol.ToString();
                                worksheet3.Cell(currentRow3, 2).Value = elem.Nazwa.ToString();
                                currentRow3++;
                            }
                        }

                        worksheet3.Range(2, 3, currentRow3 - 1, 3 + miesiace.Count).Value = 0;

                        foreach (DokumentHandlowy fv in fakturyList.Count == 0 ? faktury.ToArray<DokumentHandlowy>() : fakturyList.ToArray())
                        {
                            foreach (PozycjaDokHandlowego poz in fv.Pozycje)
                            {
                                // wyliczanie wiersza
                                e = (ElemSlownika)poz.Towar.Features["SALES"];
                                int row = dim3.IndexOf(e.Symbol.ToString()) + 2;

                                // wyliczanie kolumny
                                monthStr = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(fv.Data.Month);
                                int col = miesiace.IndexOf(monthStr.ToUpper() + " " + fv.Data.Year) + 3;

                                worksheet3.Cell(row, col).Value = (double)worksheet3.Cell(row, col).Value + (poz.CenaPoRabacie.Value * poz.Ilosc.Value * fv.KursWaluty);
                            }
                        }

                        s.Save();
                    }

                    #region total w3
                    for (int row = 2; row <= currentRow3 - 1; row++)
                        worksheet3.Cell(row, 3 + miesiace.Count).FormulaA1 = "=SUM(C" + row + ":" + worksheet3.Cell(row, 3 + miesiace.Count - 1).Address + ")";
                    for (int col = 3; col <= 3 + miesiace.Count; col++)
                        worksheet3.Cell(currentRow3, col).FormulaA1 = "=SUM(" + worksheet3.Cell(2, col).Address + ":" + worksheet3.Cell(currentRow3 - 1, col).Address + ")";
                    #endregion

                    #endregion

                    #region arkusz 4
                    using (Session s = Context.Login.CreateSession(false, false))
                    {
                        ksiegaModule = KsiegaModule.GetInstance(s);
                        elementy = ksiegaModule.ElemSlownikow.CreateView();
                        elementy.Condition &= new FieldCondition.Equal("Definicja", "Konta Norweskie");

                        foreach (ElemSlownika elem in elementy)
                        {
                            if ((bool)elem.Features["Sal"] == true)
                            {
                                dim4.Add(elem.Symbol.ToString());
                                worksheet4.Cell(currentRow4, 1).Value = elem.Symbol.ToString();
                                worksheet4.Cell(currentRow4, 2).Value = elem.Nazwa.ToString();
                                currentRow4++;
                            }
                        }

                        worksheet4.Range(2, 3, currentRow4 - 1, 3 + miesiace.Count).Value = 0;

                        foreach (DokumentHandlowy fv in fakturyList.Count == 0 ? faktury.ToArray<DokumentHandlowy>() : fakturyList.ToArray())
                        {
                            if (fv.Podmiot.Kod != "ODB-10000")
                            {
                                foreach (PozycjaDokHandlowego poz in fv.Pozycje)
                                {
                                    // wyliczanie wiersza
                                    e = (ElemSlownika)poz.Towar.Features["SALES"];
                                    int row = dim4.IndexOf(e.Symbol.ToString()) + 2;

                                    // wyliczanie kolumny
                                    monthStr = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(fv.Data.Month);
                                    int col = miesiace.IndexOf(monthStr.ToUpper() + " " + fv.Data.Year) + 3;

                                    worksheet4.Cell(row, col).Value = (double)worksheet4.Cell(row, col).Value + (poz.CenaPoRabacie.Value * poz.Ilosc.Value * fv.KursWaluty);
                                }
                            }                             
                        }

                        s.Save();
                    }

                    #region total w3
                    for (int row = 2; row <= currentRow4 - 1; row++)
                        worksheet4.Cell(row, 3 + miesiace.Count).FormulaA1 = "=SUM(C" + row + ":" + worksheet4.Cell(row, 3 + miesiace.Count - 1).Address + ")";
                    for (int col = 3; col <= 3 + miesiace.Count; col++)
                        worksheet4.Cell(currentRow4, col).FormulaA1 = "=SUM(" + worksheet4.Cell(2, col).Address + ":" + worksheet4.Cell(currentRow4 - 1, col).Address + ")";
                    #endregion

                    #endregion

                    #region headlines worksheet 1        
                    worksheet.Cell("A1").Value = "SALES INVOICES";
                    worksheet.Range(1, 1, 1, 2).Merge();
                    worksheet.Range(1, 1, 1, 2).Style.Font.FontName = "Arial";
                    worksheet.Range(1, 1, 1, 2).Style.Font.FontSize = 18;
                    worksheet.Range(1, 1, 1, 2).Style.Font.Bold = true;

                    worksheet.Cell(2, 1).Value = "Nr";
                    worksheet.Cell(2, 2).Value = "Data dokumentu";
                    worksheet.Cell(2, 3).Value = "Nr nabywcy (sprzedaż)";
                    worksheet.Cell(2, 4).Value = "Nazwa nabywcy";
                    worksheet.Cell(2, 5).Value = "Kod waluty";
                    worksheet.Cell(2, 6).Value = "Kwota netto";
                    worksheet.Cell(2, 7).Value = "Kurs waluty";
                    worksheet.Cell(2, 8).Value = "Zapłacone";
                    worksheet.Cell(2, 9).Value = "Termin płatności";
                    worksheet.Range(2, 1, 2, 9).Style.Fill.BackgroundColor = XLColor.LightGray;
                    #endregion

                    #region style worksheet 
                    for (int row = 2; row < endRow; row++)
                        for (int col = 1; col <= 9; col++)
                            worksheet.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

                    // komorka z sumą
                    worksheet.Cell(endRow, 6).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    worksheet.Cell(endRow, 6).Style.Fill.BackgroundColor = XLColor.Yellow;

                    worksheet.Row(2).Height = 26;
                    worksheet.Row(2).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet.Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Column(5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Column(6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Column(8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Column(9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Range(2, 1, 2, 6).SetAutoFilter();

                    // column number format
                    worksheet.Column(6).Style.NumberFormat.Format = "0.00";
                    worksheet.Column(13).Style.NumberFormat.Format = "0.00";

                    worksheet.SheetView.FreezeRows(2);
                    worksheet.Columns().AdjustToContents();
                    worksheet.Column(1).Width = 17;

                    worksheet.Columns(8, 9).Hide();
                    #endregion

                    #region headlines worksheet 2      
                    worksheet2.Cell("A1").Value = "SALES DETAIL";
                    worksheet2.Cell(2, 1).Value = "Nr dokumentu";
                    worksheet2.Cell(2, 2).Value = "Data księgowania";
                    worksheet2.Cell(2, 3).Value = "Klient";
                    worksheet2.Cell(2, 4).Value = "Nr artykułu";
                    worksheet2.Cell(2, 5).Value = "Nazwa artykułu";
                    worksheet2.Cell(2, 6).Value = "Jednostka";
                    worksheet2.Cell(2, 7).Value = "Ilość";
                    worksheet2.Cell(2, 8).Value = "Wartość";
                    worksheet2.Cell(2, 9).Value = "Cena";
                    worksheet2.Cell(2, 10).Value = "Waluta";
                    worksheet2.Cell(2, 11).Value = "Wartość wymiaru";
                    worksheet2.Cell(2, 12).Value = "Nazwa wymiaru";
                    worksheet2.Row(2).Style.Fill.BackgroundColor = XLColor.LightGray;
                    #endregion

                    #region style worksheet 2
                    for (int row = 2; row < currentRow2; row++)
                        for (int col = 1; col <= 24; col++)
                            worksheet2.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

                    for (int col = 1; col <= 12; col++)
                        worksheet2.Cell(2, col + 12).Value = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(col).ToString();

                    worksheet2.Columns().Style.Font.FontName = "Arial";
                    worksheet2.Columns().Style.Font.FontSize = 10;
                    worksheet2.Range(1, 1, 1, 2).Merge();
                    worksheet2.Range(1, 1, 1, 2).Style.Font.FontSize = 14;
                    worksheet2.Range(1, 1, 1, 2).Style.Font.Bold = true;

                    worksheet2.Row(2).Height = 32;
                    worksheet2.Row(2).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet2.Row(2).Style.Font.FontSize = 9;
                    worksheet2.Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet2.Range(2, 1, 2, 20).SetAutoFilter();

                    // column number format
                    worksheet2.Column(8).Style.NumberFormat.Format = "0.00";
                    worksheet2.Column(9).Style.NumberFormat.Format = "0.00";
                    for (int col = 13; col <= 24; col++)
                        worksheet2.Column(col).Style.NumberFormat.Format = "0.00";

                    worksheet2.SheetView.FreezeRows(2);
                    worksheet2.SheetView.FreezeColumns(12);
                    worksheet2.Columns().AdjustToContents();
                    worksheet2.Column(1).Width = 12;
                    worksheet2.Column(2).Width = 13;
                    worksheet2.Column(3).Width = 30;
                    worksheet2.Column(5).Width = 50;
                    #endregion

                    #region headlines worksheet 3    
                    worksheet3.Cell(1, 1).Value = "Kod";
                    worksheet3.Cell(1, 2).Value = "Dimension PLN";
                    foreach (string header in miesiace)
                    {
                        worksheet3.Cell(1, 3 + miesiace.IndexOf(header)).Value = header;
                    }
                    worksheet3.Cell(1, 3 + miesiace.Count).Value = "SUM";
                    worksheet3.Range(1, 1, 1, 3 + miesiace.Count).Style.Fill.BackgroundColor = XLColor.LightGray;
                    #endregion

                    #region style worksheet 3
                    for (int row = 1; row <= currentRow3; row++)
                        for (int col = 1; col <= 3 + miesiace.Count; col++)
                            if (!(row == currentRow3 && (col == 1 || col == 2)))
                                worksheet3.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

                    worksheet3.Range(currentRow3, 3, currentRow3, 3 + miesiace.Count).Style.Fill.BackgroundColor = XLColor.LightGray;

                    // column number format
                    worksheet3.Range(2, 3, currentRow3, 3 + miesiace.Count).Style.NumberFormat.Format = "0.00";

                    worksheet3.Range(1, 3 + miesiace.Count, currentRow3 - 1, 3 + miesiace.Count).Style.Fill.BackgroundColor = XLColor.LightGray;

                    worksheet3.Columns().AdjustToContents();
                    #endregion


                    #region headlines worksheet 4    
                    worksheet4.Cell(1, 1).Value = "Kod";
                    worksheet4.Cell(1, 2).Value = "Dimension PLN";
                    foreach (string header in miesiace)
                    {
                        worksheet4.Cell(1, 3 + miesiace.IndexOf(header)).Value = header;
                    }
                    worksheet4.Cell(1, 3 + miesiace.Count).Value = "SUM";
                    worksheet4.Range(1, 1, 1, 3 + miesiace.Count).Style.Fill.BackgroundColor = XLColor.LightGray;
                    #endregion

                    #region style worksheet 4
                    for (int row = 1; row <= currentRow4; row++)
                        for (int col = 1; col <= 3 + miesiace.Count; col++)
                            if (!(row == currentRow4 && (col == 1 || col == 2)))
                                worksheet4.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

                    worksheet4.Range(currentRow4, 3, currentRow4, 3 + miesiace.Count).Style.Fill.BackgroundColor = XLColor.LightGray;

                    // column number format
                    worksheet4.Range(2, 3, currentRow4, 3 + miesiace.Count).Style.NumberFormat.Format = "0.00";

                    worksheet4.Range(1, 3 + miesiace.Count, currentRow4 - 1, 3 + miesiace.Count).Style.Fill.BackgroundColor = XLColor.LightGray;

                    worksheet4.Columns().AdjustToContents();
                    #endregion

                    // zapis do pliku
                    workbook.SaveAs(path);
                }
            }
        }

        public bool FakturaOplacona(DokumentHandlowy fv)
        {
            if (fv.Platnosci.GetFirst() == null)
                return false;
            if (fv.BruttoCy.Value == fv.Platnosci.GetFirst().Kwota.Value)
                return true;
            return false;
        }
    }
}
