using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Charts;


namespace TestIronXL;

public static class ExampleCodeSnips
{
    public static void  ExcelComboChartExample()
    {
      // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Chart");

        // Add data which will be used by the Excel chart.
        worksheet.Cells["A1"].Value = "Name";
        worksheet.Cells["A2"].Value = "John Doe";
        worksheet.Cells["A3"].Value = "Fred Nurk";
        worksheet.Cells["A4"].Value = "Hans Meier";
        worksheet.Cells["A5"].Value = "Ivan Horvat";

        worksheet.Cells["B1"].Value = "Salary";
        worksheet.Cells["B2"].Value = 4023;
        worksheet.Cells["B3"].Value = 3263;
        worksheet.Cells["B4"].Value = 2851;
        worksheet.Cells["B5"].Value = 4694;

        worksheet.Cells["C1"].Value = "Max";
        worksheet.Cells["C2"].Value = 4500;
        worksheet.Cells["C3"].Value = 4300;
        worksheet.Cells["C4"].Value = 4000;
        worksheet.Cells["C5"].Value = 4900;

        worksheet.Cells["D1"].Value = "Min";
        worksheet.Cells["D2"].Value = 3000;
        worksheet.Cells["D3"].Value = 2800;
        worksheet.Cells["D4"].Value = 2500;
        worksheet.Cells["D5"].Value = 3400;

        // Set header row and formatting.
        worksheet.Rows[0].Style.Font.Weight = ExcelFont.BoldWeight;
        worksheet.Columns[0].SetWidth(3, LengthUnit.Centimeter);

        // Set value cells number formatting.
        foreach (var cell in worksheet.Cells.GetSubrange("B2", "D5"))
            cell.Style.NumberFormat = "\"$\"#,##0";

        // Make entire sheet print on a single page.
        worksheet.PrintOptions.FitWorksheetWidthToPages = 1;
        worksheet.PrintOptions.FitWorksheetHeightToPages = 1;

        // Create Excel combo chart and set category labels reference.
        var comboChart = worksheet.Charts.Add<ComboChart>("F2", "O25");
        comboChart.CategoryLabelsReference = "Chart!A2:A5";

        // Make chart legend visible.
        comboChart.Legend.IsVisible = true;
        comboChart.Legend.Position = ChartLegendPosition.Top;

        // Add column chart for displaying salary series.
        var salaryChart = comboChart.Add(ChartType.Column);
        salaryChart.Series.Add("=Chart!B1", "Chart!B2:B5");
        
        // Add line chart for displaying min and max series, those will use the combo chart's secondary axis.
        var minMaxChart = comboChart.Add(ChartType.Line);
        minMaxChart.Series.Add("=Chart!C1", "Chart!C2:C5");
        minMaxChart.Series.Add("=Chart!D1", "Chart!D2:D5");
        minMaxChart.UseSecondaryAxis = true;

        workbook.Save("C:\\Users\\luke\\OneDrive\\Documents\\Combo Chart.xlsx");   
    }
    
}