using ClosedXML.Excel;
using Microsoft.Win32;
using Soneta.Business;
using Soneta.Core;
using Soneta.Ksiega;
using Soneta.Types;
using System;
using System.Collections.Generic;
using System.Globalization;

[assembly: Worker(typeof(BiocontrolExcelSDK.EksportKontaNorweskie), typeof(ZapisyKsiegowe))]
[assembly: Worker(typeof(BiocontrolExcelSDK.EksportKontaNorweskie), typeof(DokEwidencja))]

namespace BiocontrolExcelSDK
{
    internal class EksportKontaNorweskie
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
            "Cost Report   ",
            Priority = 1000,
            Icon = ActionIcon.Copy,
            Mode = ActionMode.Progress,
            Target = ActionTarget.Menu | ActionTarget.ToolbarWithText)]

        public void MyActionZapisyKsiegowe()
        {
            KsiegaModule ksiegaModule;
            View elementy, zapisyKsiegowe;

            string fileName = "COST report BioControl Polska " + DateTime.Now.ToString().Remove(10) + ".xlsx";

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
                    var worksheet = workbook.Worksheets.Add("COST report BioControl Polska");

                    #region headlines
                    worksheet.Cell("A1").Value = "No.";
                    worksheet.Cell("B1").Value = "Category";
                    for (int col = 3; col <= 14; col++)
                    {
                        worksheet.Cell(1, col).Value = CultureInfo.GetCultureInfoByIetfLanguageTag("en-US").DateTimeFormat.GetMonthName(col - 2).ToString() + " " + BaseParams.Year.ToString();
                    }
                    worksheet.Cell("O1").Value = "SUM";

                    worksheet.Cell("A2").Value = "0001";
                    worksheet.Cell("B2").Value = "Sales products";
                    worksheet.Cell("B3").Value = "BioControl AS";
                    worksheet.Cell("B4").Value = "Other";

                    worksheet.Cell("A5").Value = "0002";
                    worksheet.Cell("B5").Value = "Sales services";
                    worksheet.Cell("B6").Value = "BioControl AS";
                    worksheet.Cell("B7").Value = "Other";
                    #endregion

                    int currentRow = 8, totalRows, salesRow = 666;
                    List<string> elemSymbols = new List<string>();
                    double[,] sum;

                    using (Session session = Context.Login.CreateSession(false, false))
                    {
                        ksiegaModule = KsiegaModule.GetInstance(session);

                        elementy = ksiegaModule.ElemSlownikow.CreateView();
                        elementy.Condition = new FieldCondition.Equal("Definicja", "Konta Norweskie");

                        // wypisanie elementow slownika
                        foreach (ElemSlownika es in elementy)
                        {
                            if ((bool)es.Features["Proj"] == true && es.Symbol.Length < 5)
                            {
                                worksheet.Cell(currentRow, 1).Value = es.Symbol.ToString();
                                worksheet.Cell(currentRow, 2).Value = es.Nazwa.ToString();
                                elemSymbols.Add(es.Symbol.ToString());
                                currentRow++;
                            }
                        }

                        // liczba wierszy z danymi
                        totalRows = currentRow - 2;
                        sum = new double[totalRows, 12];

                        zapisyKsiegowe = ksiegaModule.ZapisyKsiegowe.CreateView();

                        foreach (ZapisKsiegowy zapis in zapisyKsiegowe)
                        {
                            // zapisy tylko z kont zaczynających się na 4,7
                            char [] kontoCharArray = zapis.Konto.Kod.ToString().ToCharArray();
                            char kontoFirstChar = kontoCharArray[0];                           

                            if (zapis.Features["Konta Norweskie"] != null && zapis.Data.Year.ToString() == BaseParams.Year.ToString() && (kontoFirstChar == '4' || kontoFirstChar == '7'))
                            {
                                double value = (double)zapis.KwotaZapisu.Value;

                                if (zapis.Dekret.Ewidencja.Typ.ToString() == "SprzedażEwidencja" && zapis.NumerEwidencji.ToString().Contains("SPT")) // sprzedaż
                                {
                                    if (kontoCharArray[2] == '0')
                                        salesRow = 0;
                                    else if (kontoCharArray[2] == '4')
                                        salesRow = 3;
                                    else
                                        salesRow = 666;

                                    switch (salesRow)
                                    {
                                        case 666:
                                            break;
                                        default:
                                            sum[salesRow, zapis.Data.Month - 1] -= value;

                                            if (zapis.Dekret.Ewidencja.Podmiot.Nazwa.ToString() == "BioControl AS")
                                                sum[salesRow + 1, zapis.Data.Month - 1] -= value;
                                            else
                                                sum[salesRow + 2, zapis.Data.Month - 1] -= value;
                                            break;
                                    }
                                }
                                else // inne
                                {
                                    ElemSlownika es = (ElemSlownika)zapis.Features["Konta Norweskie"];

                                    if (zapis.Strona == StronaKsiegowania.Ma)
                                        value *= -1;

                                    if (elemSymbols.Contains(es.Symbol.ToString()))
                                        sum[elemSymbols.IndexOf(es.Symbol.ToString()) + 6, zapis.Data.Month - 1] += value;
                                }                             
                            }
                        }

                        session.Save();
                    }

                    #region total
                    for (int i = 0; i < sum.GetLength(0); i++)
                        for (int j = 0; j < sum.GetLength(1); j++)
                            worksheet.Cell(i + 2, j + 3).Value = sum[i, j];

                    worksheet.Cell(currentRow, 2).Value = "Gross profit";
                    worksheet.Range(currentRow, 2, currentRow, 14).Style.Fill.BackgroundColor = XLColor.LightGray;
                    worksheet.Cell(currentRow, 15).Style.Fill.BackgroundColor = XLColor.Gray;
                    worksheet.Row(currentRow).Style.Font.Bold = true;

                    // sumowanie wierszy
                    for (int row = 2; row <= currentRow - 1; row++)
                        worksheet.Cell(row, 15).FormulaA1 = "=SUM(C" + row + ":N" + row + ")";

                    // sumowanie kolumn
                    for (int col = 3; col <= 15; col++)
                        worksheet.Cell(currentRow, col).FormulaA1 = "=-SUM(" + worksheet.Cell(8, col).Address + ":" + worksheet.Cell(currentRow - 1, col).Address + ")-"
                                                                 + worksheet.Cell(2, col).Address + "-" + worksheet.Cell(5, col).Address;
                    #endregion

                    #region worksheet style
                    for (int row = 1; row <= currentRow; row++)
                        for (int col = 1; col <= 15; col++)
                            worksheet.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    worksheet.Cell(currentRow, 1).Style.Border.SetBottomBorder(XLBorderStyleValues.None);

                    worksheet.Range(1, 1, currentRow - 1, 2).Style.Fill.BackgroundColor = XLColor.LightGray;

                    worksheet.Range(1, 3, 2, 14).Style.Fill.BackgroundColor = XLColor.LightYellow;
                    worksheet.Range(5, 3, 5, 14).Style.Fill.BackgroundColor = XLColor.LightYellow;
                    worksheet.Range(1, 15, currentRow - 1, 15).Style.Fill.BackgroundColor = XLColor.Gold;

                    worksheet.Row(1).Style.Font.Bold = true;
                    worksheet.Row(1).Height = 30;
                    worksheet.Row(1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    worksheet.Column(1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    worksheet.Column(2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                    worksheet.Range(7, 1, 7, 15).Style.Border.SetBottomBorder(XLBorderStyleValues.Thick);
                    worksheet.Column(15).Style.Font.Bold = true;

                    worksheet.Rows().Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    worksheet.SheetView.FreezeRows(1);
                    worksheet.SheetView.FreezeColumns(2);
                    worksheet.Columns().AdjustToContents();

                    // column number format
                    for (int col = 3; col <= 15; col++)
                    {
                        worksheet.Column(col).Style.NumberFormat.Format = "0.00";
                        worksheet.Column(col).Width = 15;
                    }
                    #endregion

                    // zapis do pliku
                    workbook.SaveAs(path);
                }
            }           
        }
    }
}
