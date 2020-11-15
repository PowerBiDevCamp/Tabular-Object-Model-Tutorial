using System;
using Microsoft.AnalysisServices.Tabular;
using Microsoft.AnalysisServices.AdomdClient;
using System.Diagnostics;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;

class Program {

    static Server server;
    static string consoleDelimeter = "╠═══════════════════════════════════════════════════════════════════════════════╣";
    static string consoleHeader = "║";

    static string appFolder = Environment.CurrentDirectory + "\\";

    static void Main(string[] args) {

        Model model;

        Console.BackgroundColor = ConsoleColor.DarkBlue;
        Console.ForegroundColor = ConsoleColor.White;
        Console.Clear();

        server = new Server();
        if (args.Length > 0) {
            string pbiserver = args[0].ToString();
            server.Connect(pbiserver);
            model = server.Databases[0].Model;
        }
        else {
            setServer();
            model = setModel();
        }

        string userInput = "";
        while (userInput != "0") {
            Console.WriteLine(consoleDelimeter);
            myConsoleWriteLine($"Server  : {server.ConnectionString}");
            myConsoleWriteLine($"Database: {model.Database.Name}");
            Console.WriteLine(consoleDelimeter);
            myConsoleWriteLine($"    0   Exit");
            myConsoleWriteLine($"");
            myConsoleWriteLine($"    1   List Tables Storage Modes");
            myConsoleWriteLine($"    2   List Partition Queries");

            myConsoleWriteLine($"    4   List Tables");
            myConsoleWriteLine($"    5   Run DAX query");

            myConsoleWriteLine($"    7   Set Database");
            myConsoleWriteLine($"    8   Set Server");

            myConsoleWriteLine($"    9   List Table Processed state");
            myConsoleWriteLine($"   10   Get Database as TMSL");
            myConsoleWriteLine($"   11   Get Table as TMSL");
            Console.WriteLine(consoleDelimeter);


            Console.CursorVisible = true;
            Console.CursorSize = 100;     // Emphasize the cursor.
            userInput = Console.ReadLine();

            switch (userInput) {
                case "1":
                    getStorageMode(model);
                    break;
                case "2":
                    getPartitionQueries(model);
                    break;
                case "4":
                    getTables(model);
                    break;
                case "5":
                    executeDAX(model);
                    break;
                case "7":
                    model = setModel();
                    break;
                case "8":
                    setServer();
                    model = setModel();
                    break;


                case "9":
                    getTableProcessedState(model);
                    break;
                case "10":
                    getModelTMSL(model);
                    break;
                case "11":
                    getTableTMSL(model);
                    break;
                default:
                    Console.WriteLine($"You chose option: {userInput}");
                    break;
            }
        }
    }


    static void myConsoleWriteLine(string s) {
        Console.WriteLine("{0,0}  {1,-76} {0,0}", consoleHeader, s);
    }

    static private void getStorageMode(Model model) {
        foreach (Table table in model.Tables) {
            myConsoleWriteLine(
                String.Format(
                    "    {0,-30}{1,-20}",
                    table.Name,
                    table.Partitions[0].Mode
                    )
                );
        }
    }

    static private void getTableTMSL(Model model) {

        int i = 0;
        foreach (Table table in model.Tables) {
            Console.WriteLine($"{i,4} - {table.Name,-30}");
            i++;
        }

        String s = Console.ReadLine();
        int tableIndex = int.Parse(s);

        Table selectedTable = model.Tables[tableIndex];

        String tmslTable = Microsoft.AnalysisServices.Tabular.JsonSerializer.SerializeObject(selectedTable);
        String tmsl = $"{{\"createOrReplace\": {{\"object\": {{\"database\": \"{model.Database.Name}\",\"table\": \"{selectedTable.Name}\" }},\"table\": {tmslTable}}}}}";

        String tmslResultFileName = $"{appFolder}TMSLResult-{Guid.NewGuid()}.tsv";
        File.AppendAllText(tmslResultFileName, tmsl);
        ProcessStartInfo startInfo = new ProcessStartInfo();
        startInfo.FileName = "NOTEPAD.EXE";
        startInfo.Arguments = tmslResultFileName;
        Process.Start(startInfo);
    }

    static private void getPartitionQueries(Model model) {
        foreach (Table table in model.Tables) {
            switch (table.Partitions[0].SourceType) {
                case PartitionSourceType.Query:
                    QueryPartitionSource queryPartitionSource = (QueryPartitionSource)table.Partitions[0].Source;
                    myConsoleWriteLine($"Table = {table.Name} - Query =  {queryPartitionSource.Query}");
                    break;
                case PartitionSourceType.M:
                    MPartitionSource mPartitionSource = (MPartitionSource)table.Partitions[0].Source;
                    myConsoleWriteLine($"Table = {table.Name} - Expression =  {mPartitionSource.Expression}");
                    break;
            }
        }
    }

    static private void getTablesAsObjects(Model model) {
        String objectTMSL = "";
        int i = 0;
        foreach (Table table in model.Tables) {
            if (i > 0)
                objectTMSL += ",\n";

            objectTMSL += $"\t{{\n\t\"database\": \"{model.Database.Name}\",\n\t\"table\": \"{table.Name}\"\n\t}}";
            i++;
        }

        openInNotePad(objectTMSL);
    }

    static private void openInNotePad(String s) {
        String tmslResultFileName = $"{appFolder}TMSLResult-{Guid.NewGuid()}.tsv";
        File.AppendAllText(tmslResultFileName, s);
        ProcessStartInfo startInfo = new ProcessStartInfo();
        startInfo.FileName = "NOTEPAD.EXE";
        startInfo.Arguments = tmslResultFileName;
        Process.Start(startInfo);
    }

    static private void getModelTMSL(Model model) {
        String tmslDB = Microsoft.AnalysisServices.Tabular.JsonSerializer.SerializeDatabase(server.Databases[0]);
        String tmsl = $"{{\"createOrReplace\": {{\"object\": {{\"database\": \"{model.Database.Name}\"}},\"database\": {tmslDB}}}}}";
        String tmslResultFileName = $"{appFolder}TMSLResult-{Guid.NewGuid()}.tsv";
        File.AppendAllText(tmslResultFileName, tmsl);
        ProcessStartInfo startInfo = new ProcessStartInfo();
        startInfo.FileName = "NOTEPAD.EXE";
        startInfo.Arguments = tmslResultFileName;
        Process.Start(startInfo);
    }

    static private void getTables(Model model) {
        Console.WriteLine(consoleDelimeter);
        myConsoleWriteLine($"    1   Display to Screen");
        myConsoleWriteLine($"    2   Open as TMSL");
        Console.WriteLine(consoleDelimeter);
        String userInput = Console.ReadLine();
        if (userInput == "2") {
            getTablesAsObjects(model);
        }
        else {
            foreach (Table table in model.Tables) {
                myConsoleWriteLine($"  Table = {table.Name}");
            }
        }
    }

    static private void setServer() {
        string serverName = "";
        Console.Clear();
        Console.WriteLine(consoleDelimeter);
        myConsoleWriteLine("Enter or Pick a Server");
        serverName = getData("Servers.dat");
        if (server.Connected) {
            server.Disconnect();
        }
        try {
            server.Connect(serverName);
            if (!serverName.ToUpper().StartsWith("LOCALHOST")) {
                saveData(appFolder + "Servers.dat", serverName);
            }
            Console.Clear();
        }
        catch (Exception ex) {
            Console.WriteLine($"{ex.InnerException.ToString()}");
        }
    }

    static private Dictionary<int, string> setDBIndex() {

        Dictionary<int, string> databases = new Dictionary<int, string>();
        int dbIndex = 0;

        foreach (Database database in server.Databases) {
            string databaseName = database.Name;
            if (!databases.ContainsValue(databaseName)) {
                databases.Add(dbIndex, databaseName);
                dbIndex++;
            }
        }
        return databases;
    }

    static private Model setModel() {
        Model model = null;
        if (server.Databases.Count == 1) {
            model = server.Databases[0].Model;
        }
        else {
            Dictionary<int, string> x = setDBIndex();
            Console.Clear();
            Console.WriteLine(consoleDelimeter);
            myConsoleWriteLine($"Please pick a database:");
            Console.WriteLine(consoleDelimeter);
            foreach (KeyValuePair<int, string> database in x) {
                myConsoleWriteLine($"   {database.Key} - {database.Value}");
            }
            Console.WriteLine(consoleDelimeter);
            String s = Console.ReadLine();
            int dbIndex = int.Parse(s);

            if (x.ContainsKey(int.Parse(s))) {
                string db;
                x.TryGetValue(dbIndex, out db);
                model = server.Databases[db].Model;
            }
            Console.Clear();
        }
        return model;
    }

    static private void showConsoleHeader(Model model) {
        Console.WriteLine(consoleDelimeter);
        myConsoleWriteLine($"Server  : {server.ConnectionString}");
        myConsoleWriteLine($"Database: {model.Database.Name}");
        Console.WriteLine(consoleDelimeter);
    }

    static private void getTableProcessedState(Model model) {
        foreach (Table table in model.Tables) {
            foreach (Partition partition in table.Partitions) {
                String s = String.Format("{0,-30}{1,-30}{2,-20}",
                                          table.Name,
                                          partition.Name,
                                          partition.State.ToString() );
                myConsoleWriteLine(s);
            }
        }
    }

    static private string getData(String filename) {
        filename = appFolder + filename;
        Console.WriteLine(consoleDelimeter);
        if (!System.IO.File.Exists(filename)) {
            System.IO.File.AppendAllText(filename, "");
        }
        String[] lastData = System.IO.File.ReadAllLines(filename);
        if (filename.Contains("Server")) {
            List<string> p = new List<string>();
            p.AddRange(lastData);
            p.AddRange(getProcesses());
            String[] x = p.ToArray();
            lastData = p.ToArray();
        }

        int i = 0;
        foreach (string q in lastData) {
            myConsoleWriteLine($"{i}    {q}");
            i++;
        }

        Console.WriteLine(consoleDelimeter);

        String userInput = Console.ReadLine();

        if (userInput.ToUpper().StartsWith("DEL")) {
            //TODO: delete code here
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = "NOTEPAD.EXE";
            startInfo.Arguments = filename;
            Process.Start(startInfo);
        }
        else {
            String[] lines = { userInput };
            if (Int32.TryParse(userInput, out int queryNumber)) {
                userInput = lastData[queryNumber];
            }
            else {
                saveData(filename, userInput);
            }
        }
        return userInput;
    }

    public static string LookupProcess(int pid) {
        string procName;
        try { procName = Process.GetProcessById(pid).ProcessName; }
        catch (Exception) { procName = "-"; }
        return procName;
    }

    static private List<string> getProcesses() {
        Process p = new Process();
        ProcessStartInfo ps = new ProcessStartInfo();
        ps.Arguments = "-ano";
        ps.FileName = "netstat.exe";
        ps.UseShellExecute = false;
        ps.WindowStyle = ProcessWindowStyle.Hidden;
        ps.RedirectStandardInput = true;
        ps.RedirectStandardOutput = true;
        ps.RedirectStandardError = true;
        p.StartInfo = ps;
        p.Start();
        StreamReader stdOutput = p.StandardOutput;
        StreamReader stdError = p.StandardError;
        string content = stdOutput.ReadToEnd() + stdError.ReadToEnd();
        string exitStatus = p.ExitCode.ToString();
        // var Ports = new List<Port>();
        List<string> myPorts = new List<string> { };
        string[] rows = Regex.Split(content, "\r\n");
        foreach (string row in rows) {
            //Split it baby
            string[] tokens = Regex.Split(row, "\\s+");
            if (tokens.Length > 4 && (tokens[1].Equals("UDP") || tokens[1].Equals("TCP"))) {
                string localAddress = Regex.Replace(tokens[2], @"\[(.*?)\]", "1.1.1.1");
                string pname = tokens[1] == "UDP" ? LookupProcess(Convert.ToInt32(tokens[4])) : LookupProcess(Convert.ToInt32(tokens[5]));
                string pnumber = "localhost:" + localAddress.Split(':')[1];
                if (pname == "msmdsrv" && !myPorts.Contains(pnumber)) {
                    myPorts.Add(pnumber);
                }
            }
        }
        return myPorts;
    }

    static private void saveData(String filename, String line) {
        String[] lines = System.IO.File.ReadAllLines(filename);
        
        bool lineFound = false;
        foreach (string l in lines) {
            if (l == line) {
                lineFound = true;
            }
        }
        
        string[] writeThis = { line };
        if (!lineFound) {
            System.IO.File.AppendAllLines(filename, writeThis);
        }
    }


    static private void executeDAX(Model model) {
        Console.Clear();
        showConsoleHeader(model);
        myConsoleWriteLine($"Enter or Pick a Query");
        string query = getData("Query.dat");
        AdomdConnection adomdConnection = new AdomdConnection($"Data Source={model.Database.Parent.ConnectionString};Initial catalog={model.Database.Name}");
        AdomdCommand adomdCommand = new AdomdCommand(query, adomdConnection);
        adomdConnection.Open();
        String queryResultFileName = $"{appFolder}QueryResult-{Guid.NewGuid()}.tsv";
        List<string> list = new List<string>();
        bool hasHeader = false;
        try {
            AdomdDataReader reader = adomdCommand.ExecuteReader();
            while (reader.Read()) {
                String rowResults = "";
                /*****************************************************************************
                    Add Header (if needed)
                ****************************************************************************/
                if (!hasHeader) {
                    for (nt columnNumber = 0; columnNumber < reader.FieldCount; columnNumber++ ) {
                        if (columnNumber > 0) {
                            rowResults += $"\t";
                        }
                        rowResults += $"{reader.GetName(columnNumber)}";
                    }
                    Console.WriteLine(rowResults);
                    list.Add(rowResults);
                    hasHeader = true;
                }
                /*****************************************************************************
                    Add normal line
                ****************************************************************************/
                rowResults = "";
                // Create a loop for every column in the current row
                for (int columnNumber = 0; columnNumber < reader.FieldCount; columnNumber++ ) {
                    if (columnNumber > 0) {
                        rowResults += $"\t";
                    }
                    rowResults += $"{reader.GetValue(columnNumber)}";
                }
                Console.WriteLine(rowResults);
                list.Add(rowResults);
            }

            System.IO.File.WriteAllLines(queryResultFileName, list);
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
                startInfo.Arguments = queryResultFileName;
                Process.Start(startInfo);
            }

        }
        catch (Exception ex) {
            Console.WriteLine(ex.Message);
            Console.ReadKey();
        }
        adomdConnection.Close();
    }
}
