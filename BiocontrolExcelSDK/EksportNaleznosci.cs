using ClosedXML.Excel;
using Microsoft.Win32;
using Soneta.Business;
using Soneta.Kasa;
using Soneta.Types;
using System;
using System.Collections.Generic;
using System.Linq;

[assembly: Worker(typeof(BiocontrolExcelSDK.EksportNaleznosci), typeof(Naleznosc))]

namespace BiocontrolExcelSDK
{
    internal class EksportNaleznosci
    {
        public class Params : ContextBase
        {
            public Params(Context context) : base(context) { }

            [Required]
            public FromTo Okres { get; set; }
        }

        [Context]
        public Params BaseParams { get; set; }

        [Context]
        public Context Context { get; set; }

        [Action(
            "Aged Accounts Receivable",
            Priority = 30,
            Icon = ActionIcon.Copy,
            Mode = ActionMode.SingleSession,
            Target = ActionTarget.Menu | ActionTarget.ToolbarWithText)]

        public void MyActionNaleznosci()
        {
            KasaModule kasaModule;
            View naleznosci, wplaty;

            string fileName = "Aged Accounts Receivable " + DateTime.Now.ToString().Remove(10) + ".xlsx";

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
                    var worksheet = workbook.Worksheets.Add("Aged Accounts Receivable");

                    #region headlines
                    worksheet.Cell("A1").Value = "Aged Accounts Receivable";
                    worksheet.Cell("A2").Value = "BioControl Polska Spółka Z O.O";
                    worksheet.Range(2, 1, 2, 7).Merge();
                    worksheet.Cell("I5").Value = "Aged Overdue Amounts";
                    worksheet.Cells("I5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Range(5, 9, 5, 12).Merge();
                    worksheet.Range(5, 9, 5, 12).Style.Font.Bold = true;
                    worksheet.Range(5, 9, 5, 12).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    worksheet.Cell(1, 9).Value = "generated: " + DateTime.Now.ToString();
                    using (Session s = Context.Login.CreateSession(false, false))
                    {
                        worksheet.Cell(2, 9).Value = @"by: BIOCONTROL\" + s.Login.UserName.ToString();
                        s.Save();
                    }
                    worksheet.Range(1, 9, 1, 12).Merge();
                    worksheet.Range(2, 9, 2, 12).Merge();
                    worksheet.Range(1, 9, 2, 12).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    worksheet.Cell(6, 1).Value = "Posting Date";
                    worksheet.Cell(6, 2).Value = "Document Type";
                    worksheet.Cell(6, 3).Value = "Document No.";
                    worksheet.Cell(6, 4).Value = "Due Date";
                    worksheet.Cell(6, 5).Value = "Original Amount";
                    worksheet.Cell(6, 6).Value = "Balance";
                    worksheet.Cell(6, 7).Value = "Balance PLN";
                    worksheet.Cell(6, 8).Value = "Not Due";
                    worksheet.Cell(6, 9).Value = "1 - 92 days";
                    worksheet.Cell(6, 10).Value = "93 - 184 days";
                    worksheet.Cell(6, 11).Value = "185 - 275 days";
                    worksheet.Cell(6, 12).Value = "More than 275 days";
                    worksheet.Range(6, 1, 6, 12).Style.Fill.BackgroundColor = XLColor.LightGray;
                    #endregion

                    int currentRow = 8;
                    int startRow = 9;
                    string kontrOld = "", kontrNew = "";
                    string waluta = "";
                    bool firstElem = true;
                    double sumEUR = 0, sumPLN = 0, sumNOK = 0;
                    double[,] sum = new double[3, 5];
                    List<string> wypisaneRozr = new List<string>();

                    using (Session session = Context.Login.CreateSession(false, false))
                    {
                        kasaModule = KasaModule.GetInstance(session);
                        naleznosci = kasaModule.Platnosci.CreateView();
                        naleznosci.Condition &= new FieldCondition.Equal("CzyNaleznosc", true)
                                              //& new FieldCondition.Equal("Bufor", false)
                                              & new FieldCondition.Equal("Rozliczana", true)
                                              & new FieldCondition.Equal("Zrealizowane", false);
                        naleznosci.Sort = "Podmiot.Nazwa";

                        foreach (Naleznosc n in naleznosci)
                        {
                            if ((n.Dokument.Definicja.ToString() == "BOE - Bilans otwarcia" || n.Dokument.Numer.ToString().Contains("FV") || n.Dokument.Numer.ToString().Contains("KS")) && n.DataDokumentu >= BaseParams.Okres.From && n.DataDokumentu <= BaseParams.Okres.To)
                            {
                                kontrOld = kontrNew;
                                kontrNew = n.Podmiot.Nazwa.ToString();
                                var worker = new InfoPlatnoscWorker { Płatność = n };

                                #region first
                                if (firstElem)
                                {
                                    // pierwszy rekord
                                    // nowy kontrahent naglowki
                                    worksheet.Cell(currentRow, 1).Value = n.Podmiot.Kod.ToString();
                                    worksheet.Cell(currentRow, 3).Value = n.Podmiot.Nazwa.ToString();
                                    worksheet.Range(currentRow, 3, currentRow, 12).Merge();
                                    worksheet.Row(currentRow).Style.Font.Bold = true;

                                    kontrOld = kontrNew;
                                    startRow = currentRow;
                                    firstElem = false;
                                    currentRow++;
                                }
                                #endregion

                                #region other
                                if (kontrNew == kontrOld)
                                {
                                    // wypisz fakture
                                    worksheet.Cell(currentRow, 1).Value = n.DataDokumentu.ToString();
                                    worksheet.Cell(currentRow, 2).Value = n.Dokument.Definicja.ToString();
                                    worksheet.Cell(currentRow, 3).Value = n.NumerDokumentu.ToString();
                                    worksheet.Cell(currentRow, 4).Value = n.Termin.ToString();
                                    worksheet.Cell(currentRow, 5).Value = (double)n.Należność.Value;
                                    worksheet.Cell(currentRow, 6).Value = (double)n.DoRozliczenia.Value;
                                    worksheet.Cell(currentRow, 7).Value = (double)n.DoRozliczenia.Value * n.Kurs;
                                    worksheet.Range(currentRow, 8, currentRow, 12).Value = 0;
                                    worksheet.Cell(currentRow, 8 + ObliczKolumne(worker.Zwloka)).Value = (double)n.DoRozliczenia.Value;
                                }
                                else if (kontrOld != "")
                                {
                                    wplaty = kasaModule.Zaplaty.CreateView();
                                    wplaty.Condition &= new FieldCondition.Equal("Kierunek", "Przychod")
                                                        & new FieldCondition.Equal("Rozliczono", false)
                                                        & new FieldCondition.Equal("Podmiot.Nazwa", kontrOld);

                                    foreach (Zaplata z in wplaty)
                                    {
                                        if (z.Podmiot != null && z.DataDokumentu >= BaseParams.Okres.From && z.DataDokumentu <= BaseParams.Okres.To && (z.Podmiot.Kod.ToString().Contains("DOS") || z.Podmiot.Kod.ToString().Contains("ODB")))
                                        {
                                            waluta = z.Kwota.Symbol;

                                            // wypisz rekord
                                            worksheet.Cell(currentRow, 1).Value = z.DataDokumentu.ToString();
                                            worksheet.Cell(currentRow, 2).Value = "Payment";
                                            worksheet.Cell(currentRow, 3).Value = z.NumerDokumentu.ToString();
                                            if (waluta != "PLN")
                                            {
                                                worksheet.Cell(currentRow, 5).Value = (double)-z.Kwota.Value;
                                                worksheet.Cell(currentRow, 6).Value = (double)-z.DoRozliczenia.Value;
                                            }
                                            worksheet.Cell(currentRow, 7).Value = -z.Kurs * (double)z.DoRozliczenia.Value;
                                            worksheet.Range(currentRow, 8, currentRow, 12).Value = 0;

                                            #region sum
                                            switch (waluta)
                                            {
                                                case "EUR":
                                                    sumEUR -= (double)z.DoRozliczenia.Value;
                                                    sum[0, 0] -= (double)z.DoRozliczenia.Value;
                                                    break;
                                                case "PLN":
                                                    sumPLN -= (double)z.DoRozliczenia.Value;
                                                    sum[1, 0] -= (double)z.DoRozliczenia.Value;
                                                    break;
                                                case "NOK":
                                                    sumNOK -= (double)z.DoRozliczenia.Value;
                                                    sum[2, 0] -= (double)z.DoRozliczenia.Value;
                                                    break;
                                                default:
                                                    break;
                                            }
                                            #endregion

                                            wypisaneRozr.Add(z.NumerDokumentu.ToString());
                                            currentRow++;
                                        }
                                    }

                                    // suma poprzedniego kontrahenta
                                    worksheet.Cell(currentRow, 1).Value = "Total for " + kontrOld;
                                    worksheet.Range(currentRow, 1, currentRow, 4).Merge();
                                    worksheet.Cell(currentRow, 5).Value = waluta;
                                    worksheet.Cell(currentRow, 6).FormulaA1 = "=SUM(F" + startRow + ":F" + (currentRow - 1) + ")";
                                    worksheet.Cell(currentRow, 7).FormulaA1 = "=SUM(G" + startRow + ":G" + (currentRow - 1) + ")";
                                    worksheet.Cell(currentRow, 8).FormulaA1 = "=SUM(H" + startRow + ":H" + (currentRow - 1) + ")";
                                    worksheet.Cell(currentRow, 9).FormulaA1 = "=SUM(I" + startRow + ":I" + (currentRow - 1) + ")";
                                    worksheet.Cell(currentRow, 10).FormulaA1 = "=SUM(J" + startRow + ":J" + (currentRow - 1) + ")";
                                    worksheet.Cell(currentRow, 11).FormulaA1 = "=SUM(K" + startRow + ":K" + (currentRow - 1) + ")";
                                    worksheet.Cell(currentRow, 12).FormulaA1 = "=SUM(L" + startRow + ":L" + (currentRow - 1) + ")";
                                    worksheet.Row(currentRow).Style.Font.Bold = true;

                                    currentRow += 2;

                                    // nowy kontrahent naglowki
                                    worksheet.Cell(currentRow, 1).Value = n.Podmiot.Kod.ToString();
                                    worksheet.Cell(currentRow, 3).Value = n.Podmiot.Nazwa.ToString();
                                    worksheet.Range(currentRow, 3, currentRow, 12).Merge();
                                    worksheet.Row(currentRow).Style.Font.Bold = true;

                                    currentRow++;
                                    startRow = currentRow;

                                    // wypisz fakture
                                    worksheet.Cell(currentRow, 1).Value = n.DataDokumentu.ToString();
                                    worksheet.Cell(currentRow, 2).Value = n.Dokument.Definicja.ToString();
                                    worksheet.Cell(currentRow, 3).Value = n.NumerDokumentu.ToString();
                                    worksheet.Cell(currentRow, 4).Value = n.Termin.ToString();
                                    worksheet.Cell(currentRow, 5).Value = (double)n.Należność.Value;
                                    worksheet.Cell(currentRow, 6).Value = (double)n.DoRozliczenia.Value;
                                    worksheet.Cell(currentRow, 7).Value = (double)n.DoRozliczenia.Value * n.Kurs;
                                    worksheet.Range(currentRow, 8, currentRow, 12).Value = 0;
                                    worksheet.Cell(currentRow, 8 + ObliczKolumne(worker.Zwloka)).Value = (double)n.DoRozliczenia.Value;
                                }
                                #endregion

                                waluta = n.Należność.Symbol;

                                #region sum
                                if (waluta == "EUR")
                                {
                                    sumEUR += (double)n.DoRozliczenia.Value;
                                    sum[0, ObliczKolumne(worker.Zwloka)] += (double)n.DoRozliczenia.Value;
                                }
                                else if (waluta == "PLN")
                                {
                                    sumPLN += (double)n.DoRozliczenia.Value;
                                    sum[1, ObliczKolumne(worker.Zwloka)] += (double)n.DoRozliczenia.Value;
                                }
                                else if (waluta == "NOK")
                                {
                                    sumNOK += (double)n.DoRozliczenia.Value;
                                    sum[2, ObliczKolumne(worker.Zwloka)] += (double)n.DoRozliczenia.Value;
                                }
                                #endregion

                                currentRow++;
                            }
                        }

                        #region last
                        wplaty = kasaModule.Zaplaty.CreateView();
                        wplaty.Condition &= new FieldCondition.Equal("Kierunek", "Przychod")
                                            & new FieldCondition.Equal("Rozliczono", false)
                                            & new FieldCondition.Equal("Podmiot.Nazwa", kontrNew);

                        foreach (Zaplata z in wplaty)
                        {
                            if (z.Podmiot != null && z.DataDokumentu >= BaseParams.Okres.From && z.DataDokumentu <= BaseParams.Okres.To && (z.Podmiot.Kod.ToString().Contains("DOS") || z.Podmiot.Kod.ToString().Contains("ODB")))
                            {
                                waluta = z.Kwota.Symbol;

                                // wypisz rekord
                                worksheet.Cell(currentRow, 1).Value = z.DataDokumentu.ToString();
                                worksheet.Cell(currentRow, 2).Value = "Payment";
                                worksheet.Cell(currentRow, 3).Value = z.NumerDokumentu.ToString();
                                if (waluta != "PLN")
                                {
                                    worksheet.Cell(currentRow, 5).Value = (double)-z.Kwota.Value;
                                    worksheet.Cell(currentRow, 6).Value = (double)-z.DoRozliczenia.Value;
                                }
                                worksheet.Cell(currentRow, 7).Value = -z.Kurs * (double)z.DoRozliczenia.Value;
                                worksheet.Range(currentRow, 8, currentRow, 12).Value = 0;

                                #region sum
                                switch (waluta)
                                {
                                    case "EUR":
                                        sumEUR -= (double)z.DoRozliczenia.Value;
                                        sum[0, 0] -= (double)z.DoRozliczenia.Value;
                                        break;
                                    case "PLN":
                                        sumPLN -= (double)z.DoRozliczenia.Value;
                                        sum[1, 0] -= (double)z.DoRozliczenia.Value;
                                        break;
                                    case "NOK":
                                        sumNOK -= (double)z.DoRozliczenia.Value;
                                        sum[2, 0] -= (double)z.DoRozliczenia.Value;
                                        break;
                                    default:
                                        break;
                                }
                                #endregion

                                wypisaneRozr.Add(z.NumerDokumentu.ToString());
                                currentRow++;
                            }
                        }

                        // suma poprzedniego kontrahenta
                        worksheet.Cell(currentRow, 1).Value = "Total for " + kontrOld;
                        worksheet.Cell(currentRow, 5).Value = waluta;
                        worksheet.Cell(currentRow, 6).FormulaA1 = "=SUM(F" + startRow + ":F" + (currentRow - 1) + ")";
                        worksheet.Cell(currentRow, 7).FormulaA1 = "=SUM(G" + startRow + ":G" + (currentRow - 1) + ")";
                        worksheet.Cell(currentRow, 8).FormulaA1 = "=SUM(H" + startRow + ":H" + (currentRow - 1) + ")";
                        worksheet.Cell(currentRow, 9).FormulaA1 = "=SUM(I" + startRow + ":I" + (currentRow - 1) + ")";
                        worksheet.Cell(currentRow, 10).FormulaA1 = "=SUM(J" + startRow + ":J" + (currentRow - 1) + ")";
                        worksheet.Cell(currentRow, 11).FormulaA1 = "=SUM(K" + startRow + ":K" + (currentRow - 1) + ")";
                        worksheet.Cell(currentRow, 12).FormulaA1 = "=SUM(L" + startRow + ":L" + (currentRow - 1) + ")";
                        worksheet.Range(currentRow, 1, currentRow, 4).Merge();
                        worksheet.Row(currentRow).Style.Font.Bold = true;
                        #endregion

                        wplaty = kasaModule.Zaplaty.CreateView();
                        wplaty.Condition &= new FieldCondition.Equal("Kierunek", "Przychod")
                                          & new FieldCondition.Equal("Rozliczono", false);
                        wplaty.Sort = "Podmiot.Nazwa";

                        foreach (Zaplata z in wplaty)
                        {
                            if (z.Podmiot != null && !wypisaneRozr.Contains(z.NumerDokumentu.ToString()) && z.DataDokumentu >= BaseParams.Okres.From && z.DataDokumentu <= BaseParams.Okres.To && (z.Podmiot.Kod.ToString().Contains("DOS") || z.Podmiot.Kod.ToString().Contains("ODB")))
                            {
                                kontrOld = kontrNew;
                                kontrNew = z.Podmiot.Nazwa.ToString();

                                #region first
                                if (firstElem)
                                {
                                    // pierwszy rekord
                                    // nowy kontrahent naglowki
                                    worksheet.Cell(currentRow, 1).Value = z.Podmiot.Kod.ToString();
                                    worksheet.Cell(currentRow, 3).Value = z.Podmiot.Nazwa.ToString();
                                    worksheet.Range(currentRow, 3, currentRow, 12).Merge();
                                    worksheet.Row(currentRow).Style.Font.Bold = true;

                                    kontrOld = kontrNew;
                                    startRow = currentRow;
                                    firstElem = false;
                                    currentRow++;
                                }
                                #endregion

                                #region other
                                if (kontrNew == kontrOld)
                                {
                                    // wypisz rekord
                                    worksheet.Cell(currentRow, 1).Value = z.DataDokumentu.ToString();
                                    worksheet.Cell(currentRow, 2).Value = "Payment";
                                    worksheet.Cell(currentRow, 3).Value = z.NumerDokumentu.ToString();
                                    if (z.Kwota.Symbol != "PLN")
                                    {
                                        worksheet.Cell(currentRow, 5).Value = (double)-z.Kwota.Value;
                                        worksheet.Cell(currentRow, 6).Value = (double)-z.DoRozliczenia.Value;
                                    }
                                    worksheet.Cell(currentRow, 7).Value = -z.Kurs * (double)z.DoRozliczenia.Value;
                                    worksheet.Range(currentRow, 8, currentRow, 12).Value = 0;
                                }
                                else if (kontrOld != "")
                                {
                                    // suma poprzedniego kontrahenta
                                    worksheet.Cell(currentRow, 1).Value = "Total for " + kontrOld;
                                    worksheet.Range(currentRow, 1, currentRow, 4).Merge();
                                    worksheet.Cell(currentRow, 5).Value = waluta;
                                    worksheet.Cell(currentRow, 6).FormulaA1 = "=SUM(F" + startRow + ":F" + (currentRow - 1) + ")";
                                    worksheet.Cell(currentRow, 7).FormulaA1 = "=SUM(G" + startRow + ":G" + (currentRow - 1) + ")";
                                    worksheet.Cell(currentRow, 8).FormulaA1 = "=SUM(H" + startRow + ":H" + (currentRow - 1) + ")";
                                    worksheet.Cell(currentRow, 9).FormulaA1 = "=SUM(I" + startRow + ":I" + (currentRow - 1) + ")";
                                    worksheet.Cell(currentRow, 10).FormulaA1 = "=SUM(J" + startRow + ":J" + (currentRow - 1) + ")";
                                    worksheet.Cell(currentRow, 11).FormulaA1 = "=SUM(K" + startRow + ":K" + (currentRow - 1) + ")";
                                    worksheet.Cell(currentRow, 12).FormulaA1 = "=SUM(L" + startRow + ":L" + (currentRow - 1) + ")";
                                    worksheet.Row(currentRow).Style.Font.Bold = true;

                                    currentRow += 2;

                                    // nowy kontrahent naglowki
                                    worksheet.Cell(currentRow, 1).Value = z.Podmiot.Kod.ToString();
                                    worksheet.Cell(currentRow, 3).Value = z.Podmiot.Nazwa.ToString();
                                    worksheet.Range(currentRow, 3, currentRow, 12).Merge();
                                    worksheet.Row(currentRow).Style.Font.Bold = true;

                                    currentRow++;
                                    startRow = currentRow;

                                    // wypisz rekord
                                    worksheet.Cell(currentRow, 1).Value = z.DataDokumentu.ToString();
                                    worksheet.Cell(currentRow, 2).Value = "Payment";
                                    worksheet.Cell(currentRow, 3).Value = z.NumerDokumentu.ToString();
                                    if (z.Kwota.Symbol != "PLN")
                                    {
                                        worksheet.Cell(currentRow, 5).Value = (double)-z.Kwota.Value;
                                        worksheet.Cell(currentRow, 6).Value = (double)-z.DoRozliczenia.Value;
                                    }
                                    worksheet.Cell(currentRow, 7).Value = -z.Kurs * (double)z.DoRozliczenia.Value;
                                    worksheet.Range(currentRow, 8, currentRow, 12).Value = 0;
                                }
                                #endregion

                                waluta = z.Kwota.Symbol;

                                #region sum
                                switch (waluta)
                                {
                                    case "EUR":
                                        sumEUR -= (double)z.DoRozliczenia.Value;
                                        sum[0, 0] -= (double)z.DoRozliczenia.Value;
                                        break;
                                    case "PLN":
                                        sumPLN -= (double)z.DoRozliczenia.Value;
                                        sum[1, 0] -= (double)z.DoRozliczenia.Value;
                                        break;
                                    case "NOK":
                                        sumNOK -= (double)z.DoRozliczenia.Value;
                                        sum[2, 0] -= (double)z.DoRozliczenia.Value;
                                        break;
                                    default:
                                        break;
                                }
                                #endregion

                                currentRow++;
                            }
                        }

                        #region last
                        // suma poprzedniego kontrahenta
                        worksheet.Cell(currentRow, 1).Value = "Total for " + kontrOld;
                        worksheet.Cell(currentRow, 5).Value = waluta;
                        worksheet.Cell(currentRow, 6).FormulaA1 = "=SUM(F" + startRow + ":F" + (currentRow - 1) + ")";
                        worksheet.Cell(currentRow, 7).FormulaA1 = "=SUM(G" + startRow + ":G" + (currentRow - 1) + ")";
                        worksheet.Cell(currentRow, 8).FormulaA1 = "=SUM(H" + startRow + ":H" + (currentRow - 1) + ")";
                        worksheet.Cell(currentRow, 9).FormulaA1 = "=SUM(I" + startRow + ":I" + (currentRow - 1) + ")";
                        worksheet.Cell(currentRow, 10).FormulaA1 = "=SUM(J" + startRow + ":J" + (currentRow - 1) + ")";
                        worksheet.Cell(currentRow, 11).FormulaA1 = "=SUM(K" + startRow + ":K" + (currentRow - 1) + ")";
                        worksheet.Cell(currentRow, 12).FormulaA1 = "=SUM(L" + startRow + ":L" + (currentRow - 1) + ")";
                        worksheet.Range(currentRow, 1, currentRow, 4).Merge();
                        worksheet.Row(currentRow).Style.Font.Bold = true;
                        #endregion

                        session.Save();
                    }

                    currentRow += 4;

                    #region sum print
                    worksheet.Cell(currentRow, 1).Value = "Currency specification";
                    worksheet.Range(currentRow, 1, currentRow, 4).Merge();

                    worksheet.Cell(currentRow, 5).Value = "EUR";
                    worksheet.Cell(currentRow, 6).Value = sumEUR;
                    worksheet.Cell(currentRow + 1, 5).Value = "PLN";
                    worksheet.Cell(currentRow + 1, 6).Value = sumPLN;
                    worksheet.Cell(currentRow + 2, 5).Value = "NOK";
                    worksheet.Cell(currentRow + 2, 6).Value = sumNOK;

                    for (int i = 0; i < 3; i++)
                        for (int j = 0; j < 5; j++)
                            worksheet.Cell(currentRow + i, 8 + j).Value = sum[i, j];
                    #endregion

                    #region worksheet style                 
                    worksheet.Columns().Style.Font.FontName = "Arial";
                    worksheet.Columns().Style.Font.FontSize = 7;
                    worksheet.Range(1, 1, 1, 7).Merge();
                    worksheet.Range(1, 1, 1, 7).Style.Font.Bold = true;
                    worksheet.Range(1, 1, 1, 7).Style.Font.FontSize = 14;
                    worksheet.Range(currentRow, 5, currentRow + 2, 5).Style.Font.Bold = true;

                    for (int row = 1; row < currentRow; row++)
                    {
                        worksheet.Row(row).Height = 10;
                        if (row == 6)
                            for (int col = 1; col <= 12; col++)
                                worksheet.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    }

                    worksheet.Columns().Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet.Row(1).Height = 16;
                    worksheet.Row(6).Height = 16;
                    worksheet.Row(6).Style.Font.Bold = true;
                    worksheet.Column(5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    // column number format
                    worksheet.Columns(5, 12).Style.NumberFormat.Format = "0.00";

                    worksheet.SheetView.FreezeRows(6);
                    worksheet.Columns().AdjustToContents();

                    // printer settings
                    /*worksheet.PrinterSettings.RepeatRows = worksheet.Rows[1, 6].Range;
                    worksheet.HeaderFooter.OddFooter.CenteredText = "Strona " + ExcelHeaderFooter.PageNumber + " z " + ExcelHeaderFooter.NumberOfPages;
                    worksheet.PrinterSettings.Orientation = eOrientation.Landscape;
                    worksheet.PrinterSettings.FitToPage = true;
                    worksheet.PrinterSettings.FitToWidth = 1;
                    worksheet.PrinterSettings.FitToHeight = 0;
                    worksheet.PrinterSettings.PaperSize = ePaperSize.A4;*/
                    #endregion

                    // zapis do pliku
                    workbook.SaveAs(path);
                }
            }
        }

        public int ObliczKolumne(int dni)
        {
            if (dni == 0)
                return 0;
            else if (Enumerable.Range(1, 92).Contains(dni))
                return 1;
            else if (Enumerable.Range(93, 184).Contains(dni))
                return 2;
            else if (Enumerable.Range(185, 275).Contains(dni))
                return 3;
            else return 4;
        }
    }
}
