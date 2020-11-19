using System;
using Microsoft.AnalysisServices.Tabular;

namespace Learning_TOM {

  class Program {

    static void Main() {

            string DatabaseName = "Wingtip Sales MOdel";

            DatasetManager.ConnectToPowerBIAsUser();
            Database database = DatasetManager.CreateDatabase(DatabaseName);
            DatasetManager.CreateWingtipSalesModel(database);

            string newDatabaseName = "newDatabaseName";
            DatasetManager.CopyDatabase(DatabaseName, newDatabaseName);
            database = DatasetManager.CreateDatabase(newDatabaseName);

        }

    }
}
