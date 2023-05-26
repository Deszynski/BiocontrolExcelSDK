using ClosedXML.Excel;
using Microsoft.Win32;
using Soneta.Business;
using Soneta.Core;
using Soneta.Ksiega;
using Soneta.Types;
using System;
using System.Collections.Generic;
using System.Globalization;

[assembly: Worker(typeof(BiocontrolExcelSDK.EksportKontaWynikowe), typeof(Konta))]

namespace BiocontrolExcelSDK
{
    internal class EksportKontaWynikowe
    {
        public class Params : ContextBase
        {
            public Params(Context context) : base(context) { }

            public int year = DateTime.Now.Year;

            [Required]
            [Caption("Rok"), DefaultWidth(4)]
            public int Year
            {
                get => year;

                set
                {
                    year = value;
                    OnChanged(EventArgs.Empty);
                }
            }
        }

        [Context]
        public Params BaseParams { get; set; }

        [Context]
        public Context Context { get; set; }

        [Action(
            "Trial Balance by Period",
            Priority = 30,
            Icon = ActionIcon.Copy,
            Mode = ActionMode.Progress,
            Target = ActionTarget.Menu | ActionTarget.ToolbarWithText)]

        public void MyActionKonta()
        {
            KsiegaModule ksiegaModule;
            View konta, zapisy;

            string fileName = "Trial Balance by Period " + DateTime.Now.ToString().Remove(10) + ".xlsx";

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
                    var worksheet = workbook.Worksheets.Add("TrialBalanceByPeriod");

                    #region headlines
                    worksheet.Cell("A1").Value = "Bilans próbny według okresu";
                    worksheet.Cell("A2").Value = "BioControl Polska Spółka Z O.O";
                    worksheet.Range(2, 1, 2, 7).Merge();
                    worksheet.Cell(1, 8).Value = "generated: " + DateTime.Now.ToString();
                    using (Session s = Context.Login.CreateSession(false, false))
                    {
                        worksheet.Cell(2, 8).Value = @"by: BIOCONTROL\" + s.Login.UserName.ToString();
                        s.Save();
                    }
                    worksheet.Range(1, 8, 1, 14).Merge();
                    worksheet.Range(2, 8, 2, 14).Merge();
                    worksheet.Range(1, 8, 2, 14).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    for (int col = 3; col <= 14; col++)
                    {
                        worksheet.Cell(4, col).Value = CultureInfo.GetCultureInfoByIetfLanguageTag("pl-PL").DateTimeFormat.GetMonthName(col - 2).ToString() + " " + BaseParams.Year.ToString();
                    }
                    #endregion

                    int currentRow = 5;
                    List<int> symbolList = new List<int> { 400, 401, 402, 403, 404, 405, 406, 490, 700, 704, 710, 750, 755, 760, 765, 770, 771 };
                    double[] sumyMiesieczne = new double[12];

                    using (Session session = Context.Login.CreateSession(false, false))
                    {
                        ksiegaModule = KsiegaModule.GetInstance(session);

                        konta = ksiegaModule.Konta.CreateView();
                        konta.Condition &= new FieldCondition.Equal("Wynikowe", true) & new FieldCondition.Equal("Rodzaj2", "Syntetyczne");

                        foreach (KontoBase k in konta)
                        {
                            Int32.TryParse(k.Symbol.Substring(0, 3), out int symbolInt);

                            if (symbolList.Contains(symbolInt))
                            {
                                worksheet.Cell(currentRow, 1).Value = k.Symbol.ToString();
                                worksheet.Cell(currentRow, 2).Value = symbolInt != 405 ? k.Nazwa.ToString() : "Ubezp.społ.i inne świadczenia";

                                zapisy = ksiegaModule.ZapisyKsiegowe.CreateView();
                                zapisy.Condition &= new FieldCondition.Like("Konto", k.Kod.Substring(0, 3) + "*");

                                foreach (ZapisKsiegowy z in zapisy)
                                {
                                    double kwota = (double)z.KwotaZapisu.Value;

                                    if (z.Strona == StronaKsiegowania.Ma)
                                        kwota *= -1;

                                    if (z.Data.Year == BaseParams.Year)
                                        sumyMiesieczne[z.Data.Month - 1] += kwota;
                                }

                                for (int i = 0; i < 12; i++)
                                    worksheet.Cell(currentRow, i + 3).Value = sumyMiesieczne[i];

                                sumyMiesieczne = new double[12];
                                symbolList.Remove(symbolInt);
                                currentRow++;
                            }
                        }

                        session.Save();
                    }

                    #region worksheet style
                    worksheet.Columns().Style.Font.SetFontName("Calibri");
                    worksheet.Columns().Style.Font.SetFontSize(11);
                    worksheet.Range(1, 1, 1, 7).Merge();
                    worksheet.Range(1, 1, 1, 7).Style.Font.Bold = true;
                    worksheet.Range(1, 1, 1, 7).Style.Font.SetFontSize(14);

                    for (int row = 4; row < currentRow; row++)
                    {
                        if (row == 4)
                            for (int col = 3; col <= 14; col++)
                                worksheet.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                        else
                            for (int col = 1; col <= 14; col++)
                                worksheet.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    }

                    worksheet.Columns().Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet.Row(1).Height = 16;
                    worksheet.Row(4).Height = 16;
                    worksheet.Row(4).Style.Font.Bold = true;

                    // column number format
                    worksheet.Columns(3, 14).Style.NumberFormat.Format = "0.00";

                    worksheet.SheetView.FreezeRows(4);
                    worksheet.Columns().AdjustToContents();
                    #endregion

                    // zapis do pliku
                    workbook.SaveAs(path);
                }
            }
        }
    }
}
