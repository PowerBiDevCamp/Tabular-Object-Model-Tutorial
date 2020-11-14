using System;
using Microsoft.AnalysisServices.Tabular;
 
class Program {

    const string connectString = "localhost:50000"; // update with port number on your machine

    static void Main(string[] args) {
        
        Server server = new Server();
        server.Connect(connectString);
        
        Model model = server.Databases[0].Model;
        
        foreach(Table table in model.Tables) {
            Console.WriteLine($"Table : {table.Name}");
        }
    }
}
