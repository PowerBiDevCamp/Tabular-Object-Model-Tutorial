using System;
using TOM = Microsoft.AnalysisServices.Tabular;
using Microsoft.AnalysisServices.AdomdClient;
using System.Diagnostics;
using System.IO;

class Program {

    const string connectString = "Data Source=localhost:50000"; // Update port number on your machine

    static void Main(string[] args) {

        // Exercise 3 - Part 1
        ExecuteDaxQuery();

        // Exercise 3 - Part 2 
        // AddSalesRegionMeasures();
    }

    static void ExecuteDaxQuery() {

        // DAX query to be submitted totabuar database engine
        String query = @"
            EVALUATE
                SUMMARIZECOLUMNS(
                    //GROUP BY 
                    Customers[State],
                
                    //FILTER BY
                    TREATAS( {""Western Region""} , 'Customers'[Sales Region] ) ,
                     
                    // MEASURES
                    ""Sales Revenue"" , SUM(Sales[SalesAmount]) ,
                    ""Units Sold"" , SUM(Sales[Quantity])
                )
            ";

        AdomdConnection adomdConnection = new AdomdConnection(connectString);
        adomdConnection.Open();

        AdomdCommand adomdCommand = new AdomdCommand(query, adomdConnection);
        AdomdDataReader reader = adomdCommand.ExecuteReader();

        ConvertReaderToCsv(reader);

        reader.Dispose();
        adomdConnection.Close();

    }

    static void ConvertReaderToCsv(AdomdDataReader reader, bool openinExcel = true) {

        string csv = string.Empty;

        for (int col = 0; col < reader.FieldCount; col++) {
            csv += reader.GetName(col);
            csv += (col < (reader.FieldCount - 1)) ? "," : "\n";
        }

        // Create a loop for every row in the resultset
        while (reader.Read()) {
            // Create a loop for every column in the current row
            for (int i = 0; i < reader.FieldCount; i++) {
                csv += reader.GetValue(i);
                csv += (i < (reader.FieldCount - 1)) ? "," : "\n";
            }
        }

        string filePath = System.IO.Directory.GetCurrentDirectory() + @"\QueryResuts.csv";
        StreamWriter writer = File.CreateText(filePath);
        writer.Write(csv);
        writer.Flush();
        writer.Dispose();

        if (openinExcel) {
            OpenInExcel(filePath);
        }

    }

    static void OpenInExcel(string FilePath) {

        ProcessStartInfo startInfo = new ProcessStartInfo();

        bool excelFound = false;
        if (File.Exists("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE")) {
            startInfo.FileName = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE";
            excelFound = true;
        }
        else {
            if (File.Exists("C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE")) {
                startInfo.FileName = "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE";
                excelFound = true;
            }
        }
        if (excelFound) {
            startInfo.Arguments = FilePath;
            Process.Start(startInfo);
        }
        else {
            System.Console.WriteLine("Coud not find Microsoft Exce on this PC.");
        }
        
    }

    static void AddSalesRegionMeasures() {

        // DAX query to be submitted totabuar database engine
        String query = "EVALUATE( VALUES(Customers[Sales Region]) )";

        AdomdConnection adomdConnection = new AdomdConnection(connectString);
        adomdConnection.Open();

        AdomdCommand adomdCommand = new AdomdCommand(query, adomdConnection);
        AdomdDataReader reader = adomdCommand.ExecuteReader();

        // open connection use TOM to create new measures
        TOM.Server server = new TOM.Server();
        server.Connect(connectString);
        TOM.Model model = server.Databases[0].Model;
        TOM.Table salesTable = model.Tables["Sales"];

        String measureDescription = "Auto Measures";
        // delete any previously created "Auto" measures
        foreach (TOM.Measure m in salesTable.Measures) {
            if (m.Description == measureDescription) {
                salesTable.Measures.Remove(m);
                model.SaveChanges();
            }
        }

        // Create the new measures
        while (reader.Read()) {
            String SalesRegion = reader.GetValue(0).ToString();
            String measureName = $"{SalesRegion} Sales";

            TOM.Measure measure = new TOM.Measure() {
                Name = measureName,
                Description = measureDescription,
                DisplayFolder = "Auto Measures",
                FormatString = "$#,##0",                
                Expression = $@"CALCULATE( SUM(Sales[SalesAmount]), Customers[Sales Region] = ""{SalesRegion}"" )"
            };

            salesTable.Measures.Add(measure);
        }

        model.SaveChanges();
        reader.Dispose();
        adomdConnection.Close();

    }
}