using ClosedXML.Excel;
using Microsoft.Win32;
using Soneta.Business;
using Soneta.Handel;
using Soneta.Towary;
using Soneta.Types;
using System;
using System.Linq;

[assembly: Worker(typeof(BiocontrolExcelSDK.EksportDokumentyMagazynowe), typeof(DokumentHandlowy))]

namespace BiocontrolExcelSDK
{
    internal class EksportDokumentyMagazynowe
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
        public DokumentHandlowy Dokument { get; set; }

        [Context]
        public Context Context { get; set; }

        [Action(
            "Dokumenty magazynowe",
            Priority = 30,
            Icon = ActionIcon.Copy,
            Mode = ActionMode.SingleSession,
            Target = ActionTarget.Menu | ActionTarget.ToolbarWithText)]

        public void MyActionDokMag()
        {
            HandelModule handelModule;
            View dokumenty;

            string fileName = "Dokumenty magazynowe " + DateTime.Now.ToString().Remove(10) + ".xlsx";

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
                    var worksheet = workbook.Worksheets.Add("Dokumenty");

                    #region headlines
                    worksheet.Cell(1, 1).Value = "Data dokumentu";
                    worksheet.Cell(1, 2).Value = "Typ zapisu";
                    worksheet.Cell(1, 3).Value = "Nr dokumentu";
                    worksheet.Cell(1, 4).Value = "Kontrahent";
                    worksheet.Cell(1, 5).Value = "Nr zapasu";
                    worksheet.Cell(1, 6).Value = "Pełna nazwa";
                    worksheet.Cell(1, 7).Value = "Ilość";
                    worksheet.Cell(1, 8).Value = "Kwota sprzedaży\n(rzeczywista)";
                    worksheet.Cell(1, 9).Value = "Kwota kosztu\n(rzeczywista)";
                    worksheet.Cell(1, 10).Value = "Cena";
                    worksheet.Range(1, 1, 1, 10).Style.Fill.BackgroundColor = XLColor.LightGray;
                    #endregion

                    int currentRow = 2;
                    string typZapisu;
                    bool minus;
                    string[] definicje = {"PZ", "WZ", "PW", "RW"};

                    using (Session session = Context.Login.CreateSession(false, false))
                    {
                        handelModule = HandelModule.GetInstance(session);

                        dokumenty = handelModule.DokHandlowe.CreateView();
                        dokumenty.Condition &= new FieldCondition.Contain("Data", BaseParams.Okres);
                        dokumenty.Sort = "Data";

                        foreach (DokumentHandlowy dokument in dokumenty)
                        {
                            if (definicje.Any(dokument.Definicja.Symbol.Contains))
                            {
                                foreach (PozycjaDokHandlowego pozycja in dokument.Pozycje)
                                {
                                    if (pozycja.Towar.Typ != TypTowaru.Usługa)
                                    {
                                        if (dokument.Definicja.Symbol.ToString().Contains("PZ"))
                                        {
                                            typZapisu = "Zakup";
                                            minus = false;
                                        }
                                        else if (dokument.Definicja.Symbol.ToString().Contains("WZ"))
                                        {
                                            typZapisu = "Sprzedaż";
                                            minus = true;
                                        }
                                        else if (dokument.Definicja.Symbol.ToString().Contains("PW"))
                                        {
                                            typZapisu = "Przychód";
                                            minus = false;
                                        }
                                        else if (dokument.Definicja.Symbol.ToString().Contains("RW"))
                                        {
                                            typZapisu = "Rozchód";
                                            minus = true;
                                        }
                                        else
                                        {
                                            typZapisu = "";
                                            minus = false;
                                        }
                                        
                                        double ilosc = pozycja.Ilosc.Value;
                                        double wartosc = (double)pozycja.WartośćWCenieZakupu;

                                        if (minus)
                                        {
                                            ilosc *= -1;
                                            wartosc *= -1;
                                        }                                       

                                        worksheet.Cell(currentRow, 1).Value = dokument.Data.ToString();
                                        worksheet.Cell(currentRow, 2).Value = typZapisu;
                                        worksheet.Cell(currentRow, 3).Value = dokument.Numer.ToString();
                                        worksheet.Cell(currentRow, 4).Value = dokument.Kontrahent != null ? dokument.Kontrahent.Nazwa.ToString() : "";
                                        worksheet.Cell(currentRow, 5).Value = pozycja.Towar.Kod.ToString();
                                        worksheet.Cell(currentRow, 6).Value = pozycja.Towar.Nazwa.ToString();
                                        worksheet.Cell(currentRow, 7).Value = ilosc;
                                        worksheet.Cell(currentRow, 8).Value = pozycja.WartoscCy.Value;
                                        worksheet.Cell(currentRow, 9).Value = wartosc;
                                        worksheet.Cell(currentRow, 10).Value = pozycja.CenaJednostkowa.Value;

                                        currentRow++;
                                    }                                
                                }
                            }                                                  
                        }

                        session.Save();
                    }

                    #region sum
                    worksheet.Cell(currentRow, 7).Value = "SUMA";
                    worksheet.Cell(currentRow, 8).FormulaA1 = "SUM(H2:H" + (currentRow - 1) + ")";
                    #endregion

                    #region worksheet style  
                    worksheet.Range(1, 1, currentRow - 1, 10).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    worksheet.Range(1, 1, currentRow - 1, 10).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    worksheet.Cell(currentRow, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    worksheet.Cell(currentRow, 8).Style.Fill.BackgroundColor = XLColor.LightGray;

                    // column number format
                    worksheet.Columns(8, 10).Style.NumberFormat.Format = "0.00";

                    worksheet.SheetView.FreezeRows(1);
                    worksheet.Columns().AdjustToContents();
                    worksheet.Range(1, 1, 1, 6).SetAutoFilter();
                    worksheet.Row(1).Style.Alignment.WrapText = true;
                    worksheet.Row(1).Height = 50;
                    worksheet.Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Row(1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    worksheet.Column(4).Width = 50;
                    worksheet.Column(6).Width = 50;

                    #endregion

                    // zapis do pliku
                    workbook.SaveAs(path);
                }
            }
        }

        public static bool IsVisibleMyActionDokMag(DokumentHandlowy Dokument)
        {
            return Dokument.TypPartii == Soneta.Magazyny.TypPartii.Magazynowy;
        }
    }
}
