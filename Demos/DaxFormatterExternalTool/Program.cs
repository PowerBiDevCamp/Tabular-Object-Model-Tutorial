using Microsoft.AnalysisServices.Tabular;

namespace DaxFormatterExternalTool {
    
    class Program {
   
   static void Main(string[] args) {

            string connectString = args.Length >= 1 ? args[0] : "localhost:53610";

            Server server = new Server();
            server.Connect(connectString);

            Model model = server.Databases[0].Model;

            DaxFormatter.FormatDaxForModel(model);

        }

    }
}
