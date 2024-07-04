using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MySql.Data.MySqlClient;
using System.Collections.ObjectModel;
using System.Windows;
using Application = Microsoft.Office.Interop.Excel.Application;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Reflection;

namespace OcEntExport;

public partial class MainWindow : System.Windows.Window
{
    public MainWindow()
    {
        InitializeComponent();
        DataContext = new MainViewModel();
    }
}

public partial class MainViewModel : ObservableObject
{
    private Application? excelApp;

    [ObservableProperty]
    private bool _shouldOutputDateSheets = true;

    [ObservableProperty]
    private DateTime _fromDate = DateTime.Now.Date.AddDays(-1);

    [ObservableProperty]
    private DateTime _toDate = DateTime.Now.Date.AddDays(-1);

    [ObservableProperty]
    private string _progressString = "";

    public ObservableCollection<SaleViewModel> Sales { get; private set; } = new ObservableCollection<SaleViewModel>();
    public ObservableCollection<SummaryL1> SalesSummary { get; private set; } = new ObservableCollection<SummaryL1>();

    [RelayCommand]
    public async Task LoadData()
    {
        try
        {
            ProgressString = "Connecting...";

            var settingsVal = App.GetAppSettings();
            if (!settingsVal.HasValue)
            {
                MessageBox.Show("Error reading Application Settings");
                ProgressString = string.Empty;
                return;
            }

            var settings = settingsVal!.Value;
            var query = settings.Query;

            var con = new MySqlConnection($"server={settings.HostServer};user=dbadmin;password=OraS1m$1;database=CUSTOMDB;");
            try
            {
                con.Open();
            }
            catch
            {
                MessageBox.Show("Could not connect to database");
                ProgressString = string.Empty;
                return;
            }

            Sales.Clear();
            using (MySqlCommand command = new MySqlCommand(query, con))
            {
                command.Parameters.AddWithValue("@FROMDATE", FromDate.Date);
                command.Parameters.AddWithValue("@TODATE", ToDate.Date);

                using (MySqlDataReader reader = command.ExecuteReader())
                {
                    ProgressString = "Loading Data...";
                    while (await reader.ReadAsync())
                    {
                        SaleViewModel sale = new SaleViewModel
                        {
                            BusinessDate = reader["BUSINESSDATE"] != DBNull.Value ? Convert.ToDateTime(reader["BUSINESSDATE"]) : DateTime.Now.Date,
                            CheckOpen = reader["CHECKOPEN"] != DBNull.Value ? Convert.ToDateTime(reader["CHECKOPEN"]) : DateTime.Now.Date,
                            CheckClose = reader["CHECKCLOSE"] != DBNull.Value ? Convert.ToDateTime(reader["CHECKCLOSE"]) : DateTime.Now.Date,
                            CheckId = reader["CHECKID"] != DBNull.Value ? Convert.ToInt64(reader["CHECKID"]) : 0L,
                            CheckDetailId = reader["CHECKDETAILID"] != DBNull.Value ? Convert.ToInt64(reader["CHECKDETAILID"]) : 0L,
                            RoundNumber = reader["ROUNDNUMBER"] != DBNull.Value ? Convert.ToInt16(reader["ROUNDNUMBER"]) : (short)0,
                            CheckNumber = reader["CHECKNUMBER"] != DBNull.Value ? Convert.ToInt32(reader["CHECKNUMBER"]) : 0,
                            RvcNumber = reader["RVCNUMBER"] != DBNull.Value ? Convert.ToInt32(reader["RVCNUMBER"]) : 0,
                            RvcName = reader["RVCNAME"] != DBNull.Value ? reader["RVCNAME"].ToString()! : string.Empty,
                            Workstations = reader["WORKSTATIONS"] != DBNull.Value ? reader["WORKSTATIONS"].ToString()! : string.Empty,
                            OrderTypeIndex = reader["ORDERTYPEINDEX"] != DBNull.Value ? Convert.ToInt16(reader["ORDERTYPEINDEX"]) : (short)0,
                            OrderTypeName = reader["ORDERTYPENAME"] != DBNull.Value ? reader["ORDERTYPENAME"].ToString()! : string.Empty,
                            ServingPeriodNumber = reader["SERVINGPERIODNUMBER"] != DBNull.Value ? Convert.ToInt32(reader["SERVINGPERIODNUMBER"]) : 0,
                            ServingPeriodName = reader["SERVINGPERIODNAME"] != DBNull.Value ? reader["SERVINGPERIODNAME"].ToString()! : string.Empty,
                            CheckName = reader["CHECKNAME"] != DBNull.Value ? reader["CHECKNAME"].ToString()! : string.Empty,
                            DiningTableName = reader["DININGTABLENAME"] != DBNull.Value ? reader["DININGTABLENAME"].ToString()! : string.Empty,
                            Covers = reader["COVERS"] != DBNull.Value ? Convert.ToInt32(reader["COVERS"]) : 0,
                            Payment = reader["PAYMENT"] != DBNull.Value ? reader["PAYMENT"].ToString()! : string.Empty,
                            PaymentNotes = reader["PAYMENTNOTES"] != DBNull.Value ? reader["PAYMENTNOTES"].ToString()! : string.Empty,
                            MiNumber = reader["MINUMBER"] != DBNull.Value ? Convert.ToInt64(reader["MINUMBER"]) : 0L,
                            MiMasterName = reader["MIMASTERNAME"] != DBNull.Value ? reader["MIMASTERNAME"].ToString()! : string.Empty,
                            MiDefName = reader["MIDEFNAME"] != DBNull.Value ? reader["MIDEFNAME"].ToString()! : string.Empty,
                            FamilyGroupNum = reader["FAMILYGROUPNUM"] != DBNull.Value ? Convert.ToInt64(reader["FAMILYGROUPNUM"]) : 0L,
                            FamilyGroupName = reader["FAMILYGROUPNAME"] != DBNull.Value ? reader["FAMILYGROUPNAME"].ToString()! : string.Empty,
                            MajorGroupNum = reader["MAJORGROUPNUM"] != DBNull.Value ? Convert.ToInt64(reader["MAJORGROUPNUM"]) : 0L,
                            MajorGroupName = reader["MAJORGROUPNAME"] != DBNull.Value ? reader["MAJORGROUPNAME"].ToString()! : string.Empty,
                            SalesCount = reader["SALESCOUNT"] != DBNull.Value ? Convert.ToInt64(reader["SALESCOUNT"]) : 0L,
                            SalesTotal = reader["SALESTOTAL"] != DBNull.Value ? Convert.ToDecimal(reader["SALESTOTAL"]) : 0m,
                            CostTotal = reader["COSTTOTAL"] != DBNull.Value ? Convert.ToDecimal(reader["COSTTOTAL"]) : 0m,
                            DiscountTotal = reader["DISCOUNTTOTAL"] != DBNull.Value ? Convert.ToDecimal(reader["DISCOUNTTOTAL"]) : 0m,
                            SalesAfterDiscount = reader["SALESAFTERDISCOUNT"] != DBNull.Value ? Convert.ToDecimal(reader["SALESAFTERDISCOUNT"]) : 0m,
                            DiscountNames = reader["DISCOUNTNAMES"] != DBNull.Value ? reader["DISCOUNTNAMES"].ToString()! : string.Empty,
                            DiscountNotes = reader["DISCOUNTNOTES"] != DBNull.Value ? reader["DISCOUNTNOTES"].ToString()! : string.Empty,
                            Tax = reader["TAX"] != DBNull.Value ? Convert.ToDecimal(reader["TAX"]) : 0m,
                            ServiceCharge = reader["SERVICECHARGE"] != DBNull.Value ? Convert.ToDecimal(reader["SERVICECHARGE"]) : 0m,
                            MicNumber = reader["MICNUMBER"] != DBNull.Value ? Convert.ToInt32(reader["MICNUMBER"]) : 0,
                            MicName = reader["MICNAME"] != DBNull.Value ? reader["MICNAME"].ToString()! : string.Empty,
                            PrintClassNum = reader["PRINTCLASSNUM"] != DBNull.Value ? Convert.ToInt64(reader["PRINTCLASSNUM"]) : 0L,
                            PrintClassName = reader["PRINTCLASSNAME"] != DBNull.Value ? reader["PRINTCLASSNAME"].ToString()! : string.Empty,
                            EmpNumber = reader["EMPNUMBER"] != DBNull.Value ? Convert.ToInt32(reader["EMPNUMBER"]) : 0,
                            EmpName = reader["EMPNAME"] != DBNull.Value ? reader["EMPNAME"].ToString()! : string.Empty,
                            AuthEmpNumber = reader["AUTHEMPNUMBER"] != DBNull.Value ? Convert.ToInt32(reader["AUTHEMPNUMBER"]) : 0,
                            AuthEmpName = reader["AUTHEMPNAME"] != DBNull.Value ? reader["AUTHEMPNAME"].ToString()! : string.Empty
                        };

                        Sales.Add(sale);
                    }
                }
            }
            con.Close();

            await GenerateSummary();
        }
        catch(Exception ex)
        {
            if(excelApp is not null)
            {
                excelApp.Quit();
                excelApp = null;
            }

            ProgressString = "An Error Occured During Export";
            MessageBox.Show($"Exception: {ex.Message}");
            return;
        }
    }

    public async Task GenerateSummary()
    {
        ProgressString = $"Generating Summary [0/{Sales.Count}]...";
        await Task.Delay(10);
        var settingsVal = App.GetAppSettings();
        if (!settingsVal.HasValue)
        {
            MessageBox.Show("Error reading Application Settings");
            ProgressString = string.Empty;
            return;
        }
        var settings = settingsVal.Value;

        var count = 0;
        var groupedSales = Sales.Select(s => new { Sales = s, Type = GetType(s.PaymentNotes) }).GroupBy(s => s.Type, (key, grp) =>
        {
            var s1 = new SummaryL1(grp.Select(g => g.Sales.MajorGroupName).Distinct().ToList())
            {
                Type = key
            };

            var group2 = grp.Select(g => g.Sales).GroupBy(s => s.PaymentNotes, (paymentNotes, g2) =>
            {
                var s2 = new SummaryL2() { PaymentNotes = paymentNotes };

                var group3 = g2.GroupBy(s => s.RvcName, (rvcName, g3) =>
                {
                    var s3 = new SummaryL3() { RvcName = rvcName };

                    foreach (var sale in g3)
                    {
                        ProgressString = $"Generating Summary [{count++}]...";
                        s3.Result.Add(sale);
                    }

                    return s3;
                });

                foreach (var item in group3.OrderBy(i => i.RvcName))
                    s2.Result.Add(item);

                return s2;
            });

            foreach (var item in group2.OrderBy(s2 => s2.PaymentNotes))
                s1.Result.Add(item);

            return s1;
        });

        groupedSales = groupedSales.OrderBy(o => o.Type);
        SalesSummary.Clear();
        foreach (var sum in groupedSales)
        {
            SalesSummary.Add(sum);
        }

        await Export();

        string GetType(string paymentNotes)
        {
            if (!paymentNotes.Contains(" - "))
                return "N/V";
            var code = paymentNotes.Split(" - ")[0];

            if (settings.TypeMappings.Any(tm => tm.Key.StartsWith(code)))
                return settings.TypeMappings.First(tm => tm.Key.StartsWith(code)).Value;
            else
                return "N/V";
        }
    }

    public async Task Export()
    {
        ProgressString = "Exporting to Excel...";
        await Task.Delay(10);

        excelApp = new Application();
        excelApp.DisplayAlerts = false;
        var workbook = excelApp.Workbooks.Add();
        excelApp.Visible = false;

        if (ShouldOutputDateSheets)
            await GenerateDateSheets(workbook);

        var fromString = FromDate.ToString("dd-MMM-yy");
        var toString = ToDate.ToString("dd-MMM-yy");
        workbook.Title = $"OCENT from {fromString} to {toString}";

        var summarySheet = (Worksheet)workbook.Worksheets.Add();
        summarySheet.Name = "Summary";
        ProgressString = "Exporting to Excel Summary...";
        await Task.Delay(10);

        List<string> majorGroups;

        var r = 3;
        var c = 1;
        Range selection;

        setupSummaryHeaders();

        var rootStart = r;
        foreach (var s1 in SalesSummary)
        {
            var startS1Row = r;
            foreach (var s2 in s1.Result)
            {
                var startS2Row = r;
                foreach (var s3 in s2.Result)
                {
                    WriteSummary3(s1, s2, s3);
                }
                WriteSummary2(s1, s2);
                var endS2Row = r - 2;
                if (endS2Row >= startS2Row)
                {
                    selection = summarySheet.Rows[$"{startS2Row}:{endS2Row}"];
                    selection.Group();
                }
            }
            WriteSummary1(s1);
            var endS1Row = r - 2;
            if (endS1Row >= startS1Row)
            {
                selection = summarySheet.Rows[$"{startS1Row}:{endS1Row}"];
                selection.Group();
            }
        }
        WriteGrandTotals();
        var rootEnd = r - 2;
        if(rootEnd >= rootStart)
        {
            selection = summarySheet.Rows[$"{rootStart}:{rootEnd}"];
            selection.Group();
        }

        Merge(summarySheet, 1, 2, 2, 2);
        Merge(summarySheet, 1, 3, 2, 3);
        summarySheet.Columns["B:C"].AutoFit();
        summarySheet.UsedRange.Borders.LineStyle = XlLineStyle.xlContinuous;
        summarySheet.UsedRange.Borders.Weight = XlBorderWeight.xlThin;

        Worksheet extraSheet = workbook.Sheets["Sheet1"];
        if(extraSheet is not null)
            extraSheet.Delete();

        excelApp.Visible = true;
        ProgressString = "Export Completed";
        workbook.SaveAs(Filename: $"OCENT from {fromString} to {toString}");

        void setupSummaryHeaders()
        {
            int c = 1;
            summarySheet.Cells[1, c++] = "Type";
            Merge(summarySheet, 1, c - 1, 2, c - 1);
            summarySheet.Cells[1, c++] = "Payment Info";
            Merge(summarySheet, 1, c - 1, 2, c - 1);
            summarySheet.Cells[1, c++] = "Revenue Center";
            Merge(summarySheet, 1, c - 1, 2, c - 1);

            OrderMajorGroups();
            foreach (var mg in majorGroups)
            {
                summarySheet.Cells[1, c] = mg;
                Merge(summarySheet, 1, c, 1, c + 1);

                summarySheet.Cells[2, c++] = "Sales Total";
                summarySheet.Cells[2, c++] = "Cost Total";
            }

            summarySheet.Cells[1, c] = "Totals";
            Merge(summarySheet, 1, c, 1, c + 1);

            summarySheet.Cells[2, c++] = "Sales Total";
            summarySheet.Cells[2, c++] = "Cost Total";

            for (int i = 1; i <= c; i++)
            {
                summarySheet.Columns[i].ColumnWidth = 13;
            }
            FormatHeader(summarySheet.Range[summarySheet.Cells[1,1], summarySheet.Cells[2,c]]);
        }

        void OrderMajorGroups()
        {
            majorGroups = Sales.Select(s => s.MajorGroupName).Distinct().ToList();
            var temp = majorGroups.OrderBy(mg => mg).ToList();
            RemoveIfContians(temp, "Alcoholic Beverage");
            RemoveIfContians(temp, "Food");
            RemoveIfContians(temp, "Non-Alcoholic Beverage");
            majorGroups.Clear();
            majorGroups.Add("Alcoholic Beverage");
            majorGroups.Add("Food");
            majorGroups.Add("Non-Alcoholic Beverage");
            foreach (var mg in temp)
                majorGroups.Add(mg);
        }

        void RemoveIfContians(List<string> list, string majorGroup)
        {
            if(list.Contains(majorGroup))
                list.Remove(majorGroup);
        }

        void WriteSummary3(SummaryL1 s1, SummaryL2 s2, SummaryL3 s3)
        {
            summarySheet.Cells[r, c++] = s1.Type;
            summarySheet.Cells[r, c++] = s2.PaymentNotes;
            summarySheet.Cells[r, c++] = s3.RvcName;

            foreach (var majorGroupName in majorGroups)
            {
                var salesTotal = s3.GetSalesTotal(majorGroupName);
                var costTotal = s3.GetCostTotal(majorGroupName);
                summarySheet.Cells[r, c++] = salesTotal;
                summarySheet.Cells[r, c++] = costTotal;
            }

            summarySheet.Cells[r, c++] = s3.GrandSalesTotal;
            summarySheet.Cells[r, c++] = s3.GrandCostTotal;

            var s3Selection = summarySheet.Range[summarySheet.Cells[r, c - 1], summarySheet.Cells[r, 1]];
            FormatS3(s3Selection);
            r++;
            c = 1;
        }

        void WriteSummary2(SummaryL1 s1, SummaryL2 s2)
        {
            summarySheet.Cells[r, c++] = s1.Type;
            summarySheet.Cells[r, c++] = s2.PaymentNotes;
            c++;

            foreach (var majorGroupName in majorGroups)
            {
                var salesTotal = s2.GetSalesTotal(majorGroupName);
                var costTotal = s2.GetCostTotal(majorGroupName);
                summarySheet.Cells[r, c++] = salesTotal;
                summarySheet.Cells[r, c++] = costTotal;
            }

            summarySheet.Cells[r, c++] = s2.GrandSalesTotal;
            summarySheet.Cells[r, c++] = s2.GrandCostTotal;

            var s2Selection = summarySheet.Range[summarySheet.Cells[r, c - 1], summarySheet.Cells[r, 1]];
            FormatS2(s2Selection);
            r++;
            c = 1;
        }


        void WriteSummary1(SummaryL1 s1)
        {
            summarySheet.Cells[r, c++] = s1.Type;
            c++;
            c++;

            foreach (var majorGroupName in majorGroups)
            {
                var salesTotal = s1.GetSalesTotal(majorGroupName);
                var costTotal = s1.GetCostTotal(majorGroupName);
                summarySheet.Cells[r, c++] = salesTotal;
                summarySheet.Cells[r, c++] = costTotal;
            }

            summarySheet.Cells[r, c++] = s1.GrandSalesTotal;
            summarySheet.Cells[r, c++] = s1.GrandCostTotal;

            var s1Selection = summarySheet.Range[summarySheet.Cells[r, c - 1], summarySheet.Cells[r, 1]];
            FormatS1(s1Selection);
            r++;
            c = 1;
        }

        void WriteGrandTotals()
        {
            summarySheet.Cells[r, c++] = "Grand Total";
            c++;
            c++;

            foreach (var majorGroupName in majorGroups)
            {
                var salesTotal = Sales.Where(s => s.MajorGroupName == majorGroupName).Sum(s => s.SalesTotal);
                var costTotal = Sales.Where(s => s.MajorGroupName == majorGroupName).Sum(s => s.CostTotal);
                summarySheet.Cells[r, c++] = salesTotal;
                summarySheet.Cells[r, c++] = costTotal;
            }

            summarySheet.Cells[r, c++] = Sales.Sum(s => s.SalesTotal);
            summarySheet.Cells[r, c++] = Sales.Sum(s => s.CostTotal);

            var s1Selection = summarySheet.Range[summarySheet.Cells[r, c - 1], summarySheet.Cells[r, 1]];
            FormatGrandTotal(s1Selection);
            r++;
            c = 1;
        }

        void FormatHeader(Range header)
        {
            header.Interior.Pattern = XlPattern.xlPatternSolid;
            header.Interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
            header.Interior.ThemeColor = XlThemeColor.xlThemeColorDark1;
            header.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(5, 5, 20));

            header.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            header.Font.Name = "Calibri";
            header.Font.Size = 10;
            header.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
            header.Font.ThemeColor = XlThemeColor.xlThemeColorLight1;
            header.Font.ThemeFont = XlThemeFont.xlThemeFontMinor;
            header.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
        }

        void FormatS3(Range s3Selection)
        {
            s3Selection.Interior.Pattern = XlPattern.xlPatternSolid;
            s3Selection.Interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
            s3Selection.Interior.ThemeColor = XlThemeColor.xlThemeColorDark1;

            s3Selection.Font.Name = "Calibri";
            s3Selection.Font.Size = 9;
            s3Selection.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
            s3Selection.Font.ThemeColor = XlThemeColor.xlThemeColorLight1;
            s3Selection.Font.ThemeFont = XlThemeFont.xlThemeFontMinor;
        }

        void FormatS2(Range s2Selection)
        {
            s2Selection.Interior.Pattern = XlPattern.xlPatternSolid;
            s2Selection.Interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
            s2Selection.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 204));

            s2Selection.Font.Name = "Calibri";
            s2Selection.Font.Size = 9;
            s2Selection.Font.Bold = true;
            s2Selection.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
            s2Selection.Font.ThemeColor = XlThemeColor.xlThemeColorLight1;
            s2Selection.Font.ThemeFont = XlThemeFont.xlThemeFontMinor;
        }

        void FormatS1(Range s1Selection)
        {
            s1Selection.Interior.Pattern = XlPattern.xlPatternSolid;
            s1Selection.Interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
            s1Selection.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 204, 255));

            s1Selection.Font.ThemeColor = XlThemeColor.xlThemeColorLight1;
            s1Selection.Font.TintAndShade = 0;

            s1Selection.Font.Name = "Calibri";
            s1Selection.Font.Size = 9;
            s1Selection.Font.Bold = true;
            s1Selection.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
            s1Selection.Font.ThemeColor = XlThemeColor.xlThemeColorLight1;
            s1Selection.Font.ThemeFont = XlThemeFont.xlThemeFontMinor;
        }

        void FormatGrandTotal(Range selection)
        {
            selection.Interior.Pattern = XlPattern.xlPatternSolid;
            selection.Interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
            selection.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(20, 150, 255));

            selection.Font.ThemeColor = XlThemeColor.xlThemeColorLight1;
            selection.Font.TintAndShade = 0;

            selection.Font.Name = "Calibri";
            selection.Font.Size = 9;
            selection.Font.Bold = true;
            selection.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
            selection.Font.ThemeColor = XlThemeColor.xlThemeColorLight1;
            selection.Font.ThemeFont = XlThemeFont.xlThemeFontMinor;
        }
    }

    public async Task GenerateDateSheets(Workbook workbook)
    {
        foreach (var day in Sales.Select(s => s.BusinessDate).Distinct().OrderBy(day => day))
        {
            var salesOfTheDay = Sales.Where(s => s.BusinessDate == day).OrderBy(s => s.RvcNumber).ThenBy(s => s.CheckNumber).ThenBy(s => s.RoundNumber).ThenBy(s => s.CheckDetailId);
            Worksheet sheet = (Worksheet)workbook.Worksheets.Add();
            sheet.Name = day.ToString("dd-MMM-yy");
            ProgressString = $"Generating Worksheet {sheet.Name}";
            await Task.Delay(10);

            string[] columnNames = {
                "Bussiness Date", "Check Open", "Check Close", "Check Id", "Check Detail Id",
                "Round Number", "Check Number", "Revenue Center #", "Revenue Center Name", "Workstations",
                "Order Type Index", "Order Type Name", "Serving Period #", "Serving Period Name",
                "Check Name", "Dining Table Name", "Covers", "Payment", "Payment Notes",
                "Menu Item #", "Master Name", "Definition Name", "Family Group #", "Family Group Name",
                "Major Group #", "Major Group Name", "Sales Count", "Sales Total", "Cost Total",
                "Discount Total", "Sales After Discount", "Discount Names", "Discount Notes",
                "Tax", "Service Charge", "Menu Item Class #", "Menu Item Class Name", "Print Class #", "Print Class Name",
                "Employee #", "Employee Name", "Authorizing Employee #", "Authorizing Employee Name"
            };

            for (int i = 1; i <= columnNames.Length; i++)
            {
                sheet.Cells[1, i] = columnNames[i - 1];
            }
            var header = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, columnNames.Length]];
            header.Interior.Pattern = XlPattern.xlPatternSolid;
            header.Interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
            header.Interior.ThemeColor = XlThemeColor.xlThemeColorDark1;
            header.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(5, 5, 20));

            header.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            header.Font.Name = "Calibri";
            header.Font.Size = 10;
            header.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
            header.Font.ThemeColor = XlThemeColor.xlThemeColorLight1;
            header.Font.ThemeFont = XlThemeFont.xlThemeFontMinor;
            header.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

            int row = 2;
            foreach (var detail in salesOfTheDay)
            {
                sheet.Cells[row, 1] = detail.BusinessDate;
                sheet.Cells[row, 2] = detail.CheckOpen;
                sheet.Cells[row, 3] = detail.CheckClose;
                sheet.Cells[row, 4] = detail.CheckId;
                sheet.Cells[row, 5] = detail.CheckDetailId;
                sheet.Cells[row, 6] = detail.RoundNumber;
                sheet.Cells[row, 7] = detail.CheckNumber;
                sheet.Cells[row, 8] = detail.RvcNumber;
                sheet.Cells[row, 9] = detail.RvcName;
                sheet.Cells[row, 10] = detail.Workstations;
                sheet.Cells[row, 11] = detail.OrderTypeIndex;
                sheet.Cells[row, 12] = detail.OrderTypeName;
                sheet.Cells[row, 13] = detail.ServingPeriodNumber;
                sheet.Cells[row, 14] = detail.ServingPeriodName;
                sheet.Cells[row, 15] = detail.CheckName;
                sheet.Cells[row, 16] = detail.DiningTableName;
                sheet.Cells[row, 17] = detail.Covers;
                sheet.Cells[row, 18] = detail.Payment;
                sheet.Cells[row, 19] = detail.PaymentNotes;
                sheet.Cells[row, 20] = detail.MiNumber;
                sheet.Cells[row, 21] = detail.MiMasterName;
                sheet.Cells[row, 22] = detail.MiDefName;
                sheet.Cells[row, 23] = detail.FamilyGroupNum;
                sheet.Cells[row, 24] = detail.FamilyGroupName;
                sheet.Cells[row, 25] = detail.MajorGroupNum;
                sheet.Cells[row, 26] = detail.MajorGroupName;
                sheet.Cells[row, 27] = detail.SalesCount;
                sheet.Cells[row, 28] = detail.SalesTotal;
                sheet.Cells[row, 29] = detail.CostTotal;
                sheet.Cells[row, 30] = detail.DiscountTotal;
                sheet.Cells[row, 31] = detail.SalesAfterDiscount;
                sheet.Cells[row, 32] = detail.DiscountNames;
                sheet.Cells[row, 33] = detail.DiscountNotes;
                sheet.Cells[row, 34] = detail.Tax;
                sheet.Cells[row, 35] = detail.ServiceCharge;
                sheet.Cells[row, 36] = detail.MicNumber;
                sheet.Cells[row, 37] = detail.MicName;
                sheet.Cells[row, 38] = detail.PrintClassNum;
                sheet.Cells[row, 39] = detail.PrintClassName;
                sheet.Cells[row, 40] = detail.EmpNumber;
                sheet.Cells[row, 41] = detail.EmpName;
                sheet.Cells[row, 42] = detail.AuthEmpNumber;
                sheet.Cells[row, 43] = detail.AuthEmpName;
                row++;
            }

            var table = sheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, sheet.UsedRange, Missing.Value, XlYesNoGuess.xlYes, Missing.Value);
            sheet.UsedRange.Columns.AutoFit();
            var dayString = day.ToString("ddMMyyyy");
            table.Name = $"Sls_{dayString}";
        }
    }

    public void Merge(Worksheet sheet, int row1, int col1, int row2, int col2)
    {
        Range selection = sheet.Range[sheet.Cells[row1, col1], sheet.Cells[row2, col2]];
        selection.Select();
        selection.Merge();
    }
}
