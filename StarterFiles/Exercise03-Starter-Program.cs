using System;
using Microsoft.AnalysisServices.Tabular;
 
class Program {

    const string connectString = "localhost:50000"; // update for port number on your machine
  
    static void Main(string[] args) {
        
        Server server = new Server();
        server.Connect(connectString);
        Model model = server.Databases[0].Model;

        foreach (Table table in model.Tables) {
            foreach (Column column in table.Columns) {
                // determine if column is visible and numeric
                if ((column.IsHidden == false) &
                    (column.DataType == DataType.Int64 ||
                     column.DataType == DataType.Decimal ||
                     column.DataType == DataType.Double)) {

                    // add automeasure for this column new measure                      
                    string measureName = $"Sum of {column.Name} ({table.Name})";
                    string expression = $"SUM('{table.Name}'[{column.Name}])";
                    string displayFolder = "Auto Measures";

                    Measure measure = new Measure() {
                        Name = measureName,
                        Expression = expression,
                        DisplayFolder = displayFolder
                    };

                    measure.Annotations.Add(new Annotation() { Value = "This is an Auto Measure" });

                    if (!table.Measures.ContainsName(measureName)) {
                        table.Measures.Add(measure);
                    }
                    else {
                        table.Measures[measureName].Expression = expression;
                        table.Measures[measureName].DisplayFolder = displayFolder;
                    }
                }
            }
        }
        // save changes back to model in Power BI Desktop
        model.SaveChanges();
    }
  
}