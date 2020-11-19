using System;
using Microsoft.AnalysisServices.Tabular;

namespace Learning_TOM {

  class Program {

    static void Main() {

      string DatabaseName = "Demo Dataset 1";

      DatasetManager.ConnectToPowerBIAsUser();
      Database database = DatasetManager.CreateDatabase(DatabaseName);
      DatasetManager.CreateWingtipSalesModel(database);

      //string newDatabaseName = "My Cloned Dataset Copy";
      //DatasetManager.CopyDatabase(DatabaseName, newDatabaseName);
      //database = DatasetManager.CreateDatabase(newDatabaseName);

    }

  }
}
