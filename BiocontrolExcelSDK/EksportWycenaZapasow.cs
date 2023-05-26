using Microsoft.Win32;
using Soneta.Business;
using Soneta.Handel;
using Soneta.Magazyny;
using Soneta.Towary;
using Soneta.Types;
using System;
using ClosedXML.Excel;

[assembly: Worker(typeof(BiocontrolExcelSDK.EksportWycenaZapasow), typeof(Towary))]

namespace BiocontrolExcelSDK
{
    internal class EksportWycenaZapasow
    {
        public class Params : ContextBase
        {
            public Params(Context context) : base(context)
            {
                Okres = FromTo.FromEnum(DefaultListPeriod.CurrentMonth);
            }

            [Required]
            public FromTo Okres { get; set; }
        }

        [Context]
        public Params BaseParams { get; set; }

        [Context]
        public Context Context { get; set; }

        [Action(
            "Wycena zapasów",
            Priority = 30,
            Icon = ActionIcon.Copy,
            Mode = ActionMode.Progress,
            Target = ActionTarget.Menu | ActionTarget.ToolbarWithText)]

        public void MyActionZapasy()
        {
            TowaryModule towaryModule;
            HandelModule handelModule;
            View produkty, towary, pozycje;

            string fileName = "Wycena zapasów " + DateTime.Now.ToString().Remove(10) + ".xlsx";

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
                if (BaseParams.Okres.To.ToString() == "(max)")
                    BaseParams.Okres = new FromTo(BaseParams.Okres.From, Date.Now);
                if (BaseParams.Okres.From.IsNull)
                    BaseParams.Okres = new FromTo(new Date(2023, 1, 1), BaseParams.Okres.To);
                if (!(BaseParams.Okres.From.Day == 1 && BaseParams.Okres.From.Month == 1 && BaseParams.Okres.From.Year == 2023))
                    BaseParams.Okres = new FromTo(BaseParams.Okres.From.PrevDay, BaseParams.Okres.To);

                string path = saveDialog.FileName;

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Wycena zapasów");

                    #region headlines
                    worksheet.Cell("A1").Value = "Wycena zapasów";
                    worksheet.Cell("A2").Value = "BioControl Polska Spółka Z O.O";
                    worksheet.Range(2, 1, 2, 7).Merge();
                    worksheet.Cell(1, 8).Value = "generated: " + DateTime.Now.ToString();
                    using (Session s = Context.Login.CreateSession(false, false))
                    {
                        worksheet.Cell(2, 8).Value = @"by: BIOCONTROL\" + s.Login.UserName.ToString();
                        s.Save();
                    }
                    worksheet.Range(1, 8, 1, 13).Merge();
                    worksheet.Range(2, 8, 2, 13).Merge();
                    worksheet.Range(1, 8, 2, 13).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    worksheet.Cell(3, 5).Value = "Na " + BaseParams.Okres.From.ToString("dd-MM-yy");
                    worksheet.Cell(3, 7).Value = "Zwiększenia (PLN)";
                    worksheet.Cell(3, 9).Value = "Zmniejszenia (PLN)";
                    worksheet.Cell(3, 11).Value = "Na " + BaseParams.Okres.To.ToString("dd-MM-yy");

                    worksheet.Range(3, 5, 3, 13).Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Range(3, 5, 3, 6).Merge();
                    worksheet.Range(3, 5, 3, 6).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    worksheet.Range(3, 7, 3, 8).Merge();
                    worksheet.Range(3, 7, 3, 8).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    worksheet.Range(3, 9, 3, 10).Merge();
                    worksheet.Range(3, 9, 3, 10).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    worksheet.Range(3, 11, 3, 13).Merge();
                    worksheet.Range(3, 11, 3, 13).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

                    worksheet.Cell(4, 1).Value = "Nr zapasu";
                    worksheet.Cell(4, 2).Value = "Opis";
                    worksheet.Cell(4, 3).Value = "Zestawienie komponentów";
                    worksheet.Cell(4, 4).Value = "Podstawowa jednostka miary";
                    worksheet.Cell(4, 5).Value = "Ilość";
                    worksheet.Cell(4, 6).Value = "Wartość";
                    worksheet.Cell(4, 7).Value = "Ilość";
                    worksheet.Cell(4, 8).Value = "Wartość";
                    worksheet.Cell(4, 9).Value = "Ilość";
                    worksheet.Cell(4, 10).Value = "Wartość";
                    worksheet.Cell(4, 11).Value = "Ilość";
                    worksheet.Cell(4, 12).Value = "Wartość";
                    worksheet.Cell(4, 13).Value = "Koszt zaksięgowany w K/G";
                    worksheet.Range(4, 1, 4, 13).Style.Fill.BackgroundColor = XLColor.LightGray;
                    #endregion

                    int currentRow = 6, divideRow;
                    double iloscZwiekszenia = 0, iloscZmniejszenia = 0;
                    double wartoscZwiekszenia = 0, wartoscZmniejszenia = 0;

                    using (Session session = Context.Login.CreateSession(false, false))
                    {
                        towaryModule = TowaryModule.GetInstance(session);
                        handelModule = HandelModule.GetInstance(session);

                        produkty = towaryModule.Towary.CreateView();
                        produkty.Condition &= new FieldCondition.Equal("Typ", "Produkt") & new FieldCondition.Equal("Blokada", false);

                        worksheet.Cell(currentRow, 1).Value = "Produkty";
                        worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                        currentRow++;

                        foreach (Towar towar in produkty)
                        {
                            StanMagazynuWorker smwFrom = new StanMagazynuWorker
                            {
                                Towar = towar,
                                Data = BaseParams.Okres.From
                            };

                            StanMagazynuWorker smwTo = new StanMagazynuWorker
                            {
                                Towar = towar,
                                Data = BaseParams.Okres.To
                            };

                            pozycje = handelModule.PozycjeDokHan.WgTowar[towar].CreateView();

                            foreach (PozycjaDokHandlowego pozycja in pozycje)
                            {
                                if ((pozycja.Dokument.Stan == StanDokumentuHandlowego.Zatwierdzony || pozycja.Dokument.Stan == StanDokumentuHandlowego.Zablokowany) && pozycja.Data > BaseParams.Okres.From && pozycja.Data <= BaseParams.Okres.To)
                                {
                                    if (pozycja.Dokument.Definicja.Symbol.Contains("WZ") || pozycja.Dokument.Definicja.Symbol.Contains("RW")
                                        || pozycja.Dokument.Definicja.Symbol.Contains("KPLW") || pozycja.Dokument.Definicja.Symbol.Contains("KWPZ"))
                                    {
                                        iloscZmniejszenia += pozycja.Ilosc.Value;
                                        wartoscZmniejszenia += (double)pozycja.WartośćWCenieZakupu;
                                    }
                                    else if (pozycja.Dokument.Definicja.Symbol.Contains("PZ") || pozycja.Dokument.Definicja.Symbol.Contains("PW") || pozycja.Dokument.Definicja.Symbol.Contains("KPLP"))
                                    {
                                        iloscZwiekszenia += pozycja.Ilosc.Value;
                                        wartoscZwiekszenia += (double)pozycja.WartośćWCenieZakupu;
                                    }
                                }
                            }

                            if (!(smwFrom.StanMagazynu.Value == 0 && smwTo.StanMagazynu.Value == 0 && iloscZwiekszenia == 0 && iloscZmniejszenia == 0))
                            {
                                worksheet.Cell(currentRow, 1).Value = towar.Kod.ToString();
                                worksheet.Cell(currentRow, 2).Value = towar.Nazwa.ToString();
                                worksheet.Cell(currentRow, 3).Value = towar.ElementyKompletu.Count > 1 ? "Tak" : "Nie";
                                worksheet.Cell(currentRow, 4).Value = towar.Jednostka.ToString();
                                worksheet.Cell(currentRow, 5).Value = smwFrom.StanMagazynu.Value.ToString();
                                worksheet.Cell(currentRow, 6).Value = smwFrom.WartośćMagazynu;
                                worksheet.Cell(currentRow, 7).Value = iloscZwiekszenia.ToString();
                                worksheet.Cell(currentRow, 8).Value = wartoscZwiekszenia;
                                worksheet.Cell(currentRow, 9).Value = iloscZmniejszenia.ToString();
                                worksheet.Cell(currentRow, 10).Value = wartoscZmniejszenia;
                                worksheet.Cell(currentRow, 11).Value = smwTo.StanMagazynu.Value.ToString();
                                worksheet.Cell(currentRow, 12).Value = smwTo.WartośćMagazynu;
                                worksheet.Cell(currentRow, 13).Value = smwTo.WartośćKsięgowaMagazynu;

                                iloscZwiekszenia = 0;
                                wartoscZwiekszenia = 0;
                                iloscZmniejszenia = 0;
                                wartoscZmniejszenia = 0;

                                currentRow++;
                            }
                        }

                        worksheet.Cell(currentRow, 1).Value = "Suma produkty";
                        worksheet.Cell(currentRow, 6).FormulaA1 = "SUM(F7:F" + (currentRow - 1) + ")";
                        worksheet.Cell(currentRow, 8).FormulaA1 = "SUM(H7:H" + (currentRow - 1) + ")";
                        worksheet.Cell(currentRow, 10).FormulaA1 = "SUM(J7:J" + (currentRow - 1) + ")";
                        worksheet.Cell(currentRow, 12).FormulaA1 = "SUM(L7:L" + (currentRow - 1) + ")";
                        worksheet.Cell(currentRow, 13).FormulaA1 = "SUM(M7:M" + (currentRow - 1) + ")";
                        worksheet.Row(currentRow).Style.Font.Bold = true;
                        worksheet.Range(currentRow, 1, currentRow, 13).Style.Fill.BackgroundColor = XLColor.LightGray;


                        for (int row = 7; row < currentRow; row++)
                            for (int col = 1; col <= 13; col++)
                                worksheet.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

                        currentRow += 2;

                        towary = towaryModule.Towary.CreateView();
                        towary.Condition &= new FieldCondition.Equal("Typ", "Towar") & new FieldCondition.Equal("Blokada", false);

                        worksheet.Cell(currentRow, 1).Value = "Towary";
                        worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                        currentRow++;

                        divideRow = currentRow;

                        foreach (Towar towar in towary)
                        {
                            StanMagazynuWorker smwFrom = new StanMagazynuWorker
                            {
                                Towar = towar,
                                Data = BaseParams.Okres.From
                            };

                            StanMagazynuWorker smwTo = new StanMagazynuWorker
                            {
                                Towar = towar,
                                Data = BaseParams.Okres.To
                            };

                            pozycje = handelModule.PozycjeDokHan.WgTowar[towar].CreateView();

                            foreach (PozycjaDokHandlowego pozycja in pozycje)
                            {
                                if ((pozycja.Dokument.Stan == StanDokumentuHandlowego.Zatwierdzony || pozycja.Dokument.Stan == StanDokumentuHandlowego.Zablokowany) && pozycja.Data > BaseParams.Okres.From && pozycja.Data <= BaseParams.Okres.To)
                                {
                                    if (pozycja.Dokument.Definicja.Symbol.Contains("WZ") || pozycja.Dokument.Definicja.Symbol.Contains("RW")
                                     || pozycja.Dokument.Definicja.Symbol.Contains("KPLW") || pozycja.Dokument.Definicja.Symbol.Contains("KWPZ"))
                                    {
                                        iloscZmniejszenia += pozycja.Ilosc.Value;
                                        wartoscZmniejszenia += (double)pozycja.WartośćWCenieZakupu;
                                    }
                                    else if (pozycja.Dokument.Definicja.Symbol.Contains("PZ") || pozycja.Dokument.Definicja.Symbol.Contains("PW") || pozycja.Dokument.Definicja.Symbol.Contains("KPLP"))
                                    {
                                        iloscZwiekszenia += pozycja.Ilosc.Value;
                                        wartoscZwiekszenia += (double)pozycja.WartośćWCenieZakupu;
                                    }
                                }
                            }

                            if (!(smwFrom.StanMagazynu.Value == 0 && smwTo.StanMagazynu.Value == 0 && iloscZwiekszenia == 0 && iloscZmniejszenia == 0))
                            {
                                worksheet.Cell(currentRow, 1).Value = towar.Kod.ToString();
                                worksheet.Cell(currentRow, 2).Value = towar.Nazwa.ToString();
                                worksheet.Cell(currentRow, 3).Value = towar.ElementyKompletu.Count > 1 ? "Tak" : "Nie";
                                worksheet.Cell(currentRow, 4).Value = towar.Jednostka.ToString();
                                worksheet.Cell(currentRow, 5).Value = smwFrom.StanMagazynu.Value.ToString();
                                worksheet.Cell(currentRow, 6).Value = smwFrom.WartośćMagazynu;
                                worksheet.Cell(currentRow, 7).Value = iloscZwiekszenia.ToString();
                                worksheet.Cell(currentRow, 8).Value = wartoscZwiekszenia;
                                worksheet.Cell(currentRow, 9).Value = iloscZmniejszenia.ToString();
                                worksheet.Cell(currentRow, 10).Value = wartoscZmniejszenia;
                                worksheet.Cell(currentRow, 11).Value = smwTo.StanMagazynu.Value.ToString();
                                worksheet.Cell(currentRow, 12).Value = smwTo.WartośćMagazynu;
                                worksheet.Cell(currentRow, 13).Value = smwTo.WartośćKsięgowaMagazynu;

                                iloscZwiekszenia = 0;
                                wartoscZwiekszenia = 0;
                                iloscZmniejszenia = 0;
                                wartoscZmniejszenia = 0;

                                currentRow++;
                            }
                        }

                        session.Save();
                    }

                    worksheet.Cell(currentRow, 1).Value = "Suma towary";
                    worksheet.Cell(currentRow, 6).FormulaA1 = "SUM(F" + divideRow + ":F" + (currentRow - 1) + ")";
                    worksheet.Cell(currentRow, 8).FormulaA1 = "SUM(H" + divideRow + ":H" + (currentRow - 1) + ")";
                    worksheet.Cell(currentRow, 10).FormulaA1 = "SUM(J" + divideRow + ":J" + (currentRow - 1) + ")";
                    worksheet.Cell(currentRow, 12).FormulaA1 = "SUM(L" + divideRow + ":L" + (currentRow - 1) + ")";
                    worksheet.Cell(currentRow, 13).FormulaA1 = "SUM(M" + divideRow + ":M" + (currentRow - 1) + ")";
                    worksheet.Row(currentRow).Style.Font.Bold = true;
                    worksheet.Range(currentRow, 1, currentRow, 13).Style.Fill.BackgroundColor = XLColor.LightGray;

                    #region sum
                    worksheet.Cell(currentRow + 3, 1).Value = "SUMA";
                    worksheet.Cell(currentRow + 3, 6).FormulaA1 = worksheet.Cell(divideRow - 3, 6).Address + "+" + worksheet.Cell(currentRow, 6).Address;
                    worksheet.Cell(currentRow + 3, 8).FormulaA1 = worksheet.Cell(divideRow - 3, 8).Address + "+" + worksheet.Cell(currentRow, 8).Address;
                    worksheet.Cell(currentRow + 3, 10).FormulaA1 = worksheet.Cell(divideRow - 3, 10).Address + "+" + worksheet.Cell(currentRow, 10).Address;
                    worksheet.Cell(currentRow + 3, 12).FormulaA1 = worksheet.Cell(divideRow - 3, 12).Address + "+" + worksheet.Cell(currentRow, 12).Address;
                    worksheet.Cell(currentRow + 3, 13).FormulaA1 = worksheet.Cell(divideRow - 3, 13).Address + "+" + worksheet.Cell(currentRow, 13).Address;
                    worksheet.Row(currentRow + 3).Style.Font.Bold = true;
                    worksheet.Range(currentRow + 3, 1, currentRow + 3, 13).Style.Fill.BackgroundColor = XLColor.LightGray;
                    #endregion

                    #region worksheet style
                    worksheet.Columns().Style.Font.SetFontName("Arial");
                    worksheet.Columns().Style.Font.SetFontSize(8);
                    worksheet.Range(1, 1, 1, 7).Merge();
                    worksheet.Range(1, 1, 1, 7).Style.Font.Bold = true;
                    worksheet.Range(1, 1, 1, 7).Style.Font.SetFontSize(14);

                    for (int col = 1; col <= 13; col++)
                        worksheet.Cell(4, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

                    for (int row = divideRow; row < currentRow; row++)
                        for (int col = 1; col <= 13; col++)
                            worksheet.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

                    worksheet.Range(6, 1, 6, 13).Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Range(divideRow - 1, 1, divideRow - 1, 13).Style.Fill.BackgroundColor = XLColor.LightGray;

                    worksheet.Columns().Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet.Columns(5, 13).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Row(3).Style.Font.Bold = true;
                    worksheet.Row(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Row(4).Style.Font.Bold = true;
                    worksheet.Row(4).Style.Alignment.WrapText = true;

                    // column number format
                    worksheet.Columns(5, 13).Style.NumberFormat.Format = "0.00";

                    worksheet.SheetView.FreezeRows(4);
                    worksheet.Columns().AdjustToContents();
                    worksheet.Column(2).Width = 50;

                    // print settings
                    worksheet.PageSetup.SetRowsToRepeatAtTop(1, 4);
                    worksheet.PageSetup.Footer.Center.AddText("Strona ", XLHFOccurrence.AllPages);
                    worksheet.PageSetup.Footer.Center.AddText(XLHFPredefinedText.PageNumber, XLHFOccurrence.AllPages);
                    worksheet.PageSetup.Footer.Center.AddText(" z ", XLHFOccurrence.AllPages);
                    worksheet.PageSetup.Footer.Center.AddText(XLHFPredefinedText.NumberOfPages, XLHFOccurrence.AllPages);
                    worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
                    worksheet.PageSetup.PagesWide = 1;
                    worksheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
                    worksheet.PageSetup.Margins.SetLeft(0.8);
                    worksheet.PageSetup.Margins.SetRight(0.8);
                    worksheet.PageSetup.Margins.SetTop(0.8);
                    worksheet.PageSetup.Margins.SetBottom(0.8);
                    worksheet.PageSetup.Margins.SetHeader(0.3);
                    worksheet.PageSetup.Margins.SetFooter(0.3);
                    #endregion

                    // zapis do pliku
                    workbook.SaveAs(path);
                }
            }                  
        }
    }
}
