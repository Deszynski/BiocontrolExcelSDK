using ClosedXML.Excel;
using Microsoft.Win32;
using Soneta.Business;
using Soneta.Ksiega;
using Soneta.Types;
using System;
using System.Collections.Generic;

[assembly: Worker(typeof(BiocontrolExcelSDK.EksportObrotySaldaKontKG), typeof(Konto))]

namespace BiocontrolExcelSDK
{
    internal class EksportObrotySaldaKontKG
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
            "Zestawienie obrótów i sald kont KG",
            Priority = 30,
            Icon = ActionIcon.Copy,
            Mode = ActionMode.Progress,
            Target = ActionTarget.Menu | ActionTarget.ToolbarWithText)]

        public void MyActionKonta()
        {
            KsiegaModule ksiegaModule;
            View kontaWszystkie, kontaSyntetyczne;

            string fileName = "Zest obrotów i sald kont KG " + DateTime.Now.ToString().Remove(10) + ".xlsx";

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
                    var worksheet = workbook.Worksheets.Add("Zest obrotów i sald kont KG");

                    #region headlines
                    worksheet.Cell("A1").Value = "Zestawienie obrotów i sald kont KG";
                    worksheet.Cell("A2").Value = "BioControl Polska Spółka Z O.O";
                    worksheet.Range(2, 1, 2, 7).Merge();
                    worksheet.Cell(1, 8).Value = "generated: " + DateTime.Now.ToString();
                    using (Session s = Context.Login.CreateSession(false, false))
                    {
                        worksheet.Cell(2, 8).Value = @"by: BIOCONTROL\" + s.Login.UserName.ToString();
                        s.Save();
                    }
                    worksheet.Range(1, 8, 1, 12).Merge();
                    worksheet.Range(2, 8, 2, 12).Merge();
                    worksheet.Range(1, 8, 2, 12).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    worksheet.Cell(4, 3).Value = "Bilans otwarcia";
                    worksheet.Range(4, 3, 4, 4).Merge();
                    worksheet.Cell(4, 5).Value = "Obroty okresu";
                    worksheet.Range(4, 5, 4, 6).Merge();
                    worksheet.Cell(4, 7).Value = "Saldo okresu";
                    worksheet.Range(4, 7, 4, 8).Merge();
                    worksheet.Cell(4, 9).Value = "Obroty narastająco";
                    worksheet.Range(4, 9, 4, 10).Merge();
                    worksheet.Cell(4, 11).Value = "Saldo na dzień";
                    worksheet.Range(4, 11, 4, 12).Merge();

                    worksheet.Cell(5, 1).Value = "Nr";
                    worksheet.Cell(5, 2).Value = "Nazwa";
                    worksheet.Cell(5, 3).Value = "Debet";
                    worksheet.Cell(5, 4).Value = "Kredyt";
                    worksheet.Cell(5, 5).Value = "Debet";
                    worksheet.Cell(5, 6).Value = "Kredyt";
                    worksheet.Cell(5, 7).Value = "Debet";
                    worksheet.Cell(5, 8).Value = "Kredyt";
                    worksheet.Cell(5, 9).Value = "Debet";
                    worksheet.Cell(5, 10).Value = "Kredyt";
                    worksheet.Cell(5, 11).Value = "Debet";
                    worksheet.Cell(5, 12).Value = "Kredyt";
                    #endregion

                    int currentRow = 6;
                    List<string> wypisane = new List<string>();

                    using (Session session = Context.Login.CreateSession(false, false))
                    {
                        ksiegaModule = KsiegaModule.GetInstance(session);

                        kontaWszystkie = ksiegaModule.Konta.CreateView();
                        kontaWszystkie.Sort = "Kod";

                        kontaSyntetyczne = ksiegaModule.Konta.CreateView();
                        kontaSyntetyczne.Condition &= new FieldCondition.Equal("Rodzaj2", "Syntetyczne");

                        foreach (KontoBase glowne in kontaSyntetyczne)
                        {
                            if (glowne.FirstChangeInfo.ToString().Contains(DateTime.Now.Year.ToString()))
                                foreach (KontoBase podrzedne in glowne.SubKonta)
                                {
                                    if (!wypisane.Contains(podrzedne.Kod))
                                    {
                                        // wypisz konto podrzedne
                                        worksheet.Cell(currentRow, 1).Value = podrzedne.Kod;
                                        worksheet.Cell(currentRow, 2).Value = podrzedne.Nazwa;

                                        ObrotyKontaWorker worker = new ObrotyKontaWorker
                                        {
                                            Param = new ObrotyKontaWorker.Params
                                            {
                                                Typ = TypObrotu.Księgowy,
                                                Bufor = true,
                                                Okres = BaseParams.Okres,
                                                NarastajacoZBO = false
                                            },
                                            Konto = podrzedne
                                        };

                                        worksheet.Cell(currentRow, 3).Value = worker.SaldoBOWn;
                                        worksheet.Cell(currentRow, 4).Value = worker.SaldoBOMa;
                                        worksheet.Cell(currentRow, 5).Value = worker.ObrotyWn;
                                        worksheet.Cell(currentRow, 6).Value = worker.ObrotyMa;
                                        worksheet.Cell(currentRow, 7).Value = worker.SaldoWn;
                                        worksheet.Cell(currentRow, 8).Value = worker.SaldoMa;
                                        worksheet.Cell(currentRow, 9).Value = worker.ObrotyNWn;
                                        worksheet.Cell(currentRow, 10).Value = worker.ObrotyNMa;

                                        if (worker.PerSaldo >= 0)
                                        {
                                            worksheet.Cell(currentRow, 11).Value = worker.PerSaldo;
                                            worksheet.Cell(currentRow, 12).Value = default(double);
                                        }
                                        else
                                        {
                                            worksheet.Cell(currentRow, 11).Value = default(double);
                                            worksheet.Cell(currentRow, 12).Value = worker.PerSaldo * (-1);
                                        }

                                        wypisane.Add(podrzedne.Kod);
                                        currentRow++;
                                    }
                                }
                        }

                        session.Save();
                    }

                    #region sum
                    worksheet.Cell(currentRow + 1, 2).Value = "Suma";
                    for (int col = 3; col <= 12; col++)
                        worksheet.Cell(currentRow + 1, col).FormulaA1 = "SUM(" + worksheet.Cell(6, col).Address + ":" + worksheet.Cell(currentRow - 1, col).Address + ")";

                    worksheet.Range(currentRow + 1, 2, currentRow + 1, 12).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    #endregion

                    #region worksheet style
                    worksheet.Columns().Style.Font.SetFontName("Calibri");
                    worksheet.Columns().Style.Font.SetFontSize(11);

                    worksheet.Range(1, 1, 1, 7).Merge();
                    worksheet.Range(1, 1, 1, 7).Style.Font.SetBold();
                    worksheet.Range(1, 1, 1, 7).Style.Font.SetFontSize(14);

                    for (int row = 4; row < currentRow; row++)
                    {
                        if (row == 4)
                            for (int col = 3; col <= 12; col++)
                                worksheet.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                        else
                            for (int col = 1; col <= 12; col++)
                                worksheet.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    }

                    worksheet.Columns().Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet.Row(1).Height = 16;
                    worksheet.Row(4).Height = 16;
                    worksheet.Row(5).Height = 16;
                    worksheet.Row(4).Style.Font.Bold = true;
                    worksheet.Row(5).Style.Font.Bold = true;

                    // column number format
                    worksheet.Columns(3, 12).Style.NumberFormat.Format = "0.00";

                    worksheet.SheetView.FreezeRows(5);
                    worksheet.Columns().AdjustToContents();
                    #endregion

                    // zapis do pliku
                    workbook.SaveAs(path);
                }               
            }
        }
    }
}
