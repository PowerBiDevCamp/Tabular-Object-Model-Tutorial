using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using Newtonsoft.Json;
using Microsoft.AnalysisServices.Tabular;
using System.Security.Cryptography;

namespace DaxFormatterExternalTool
{

    class DaxFormatter
    {

        private const string restUrl = "https://daxformatter.azurewebsites.net/api/daxformatter/DaxTextFormat";

        public static void FormatDaxForModel(Model model)
        {

            Console.WriteLine("Iterating measure, calculated columns and calulated tables for DAX formatting...");
            Console.WriteLine();

            foreach (Table table in model.Tables)
            {

                // check which measures require formatting
                foreach (var measure in table.Measures)
                {
                    Console.Write("Measure: " + table.Name + "[" + measure.Name + "]");
                    if (MeasureRequiresFormatting(measure))
                    {
                        Console.WriteLine(" - formatting DAX...");
                        string expressionOwner = measure.Name;
                        string originalDaxExpression = measure.Expression;
                        string formattedDaxExpression = DaxFormatter.FormatDaxExpression(originalDaxExpression, expressionOwner);
                        // write formatted expression back to measure
                        measure.Expression = formattedDaxExpression;
                        // store hash of formatted expression for later comparison
                        string hashedDaxExpression = GetHashValueAsString(formattedDaxExpression);
                        if (measure.Annotations.Contains("HashedExpression"))
                        {
                            measure.Annotations["HashedExpression"].Value = hashedDaxExpression;
                        }
                        else
                        {
                            measure.Annotations.Add(new Annotation { Name = "HashedExpression", Value = hashedDaxExpression });
                        }
                    }
                    else
                    {
                        Console.WriteLine(" - DAX already formatted");
                    }
                }

                // check which calculated columns require formatting
                foreach (var column in table.Columns)
                {
                    if (column.Type == ColumnType.Calculated)
                    {
                        CalculatedColumn col = (CalculatedColumn)column;
                        Console.Write("Calculated column: " + table.Name + "[" + col.Name + "]");
                        if (CalculatedColumnRequiresFormatting(col))
                        {
                            Console.WriteLine(" - formatting DAX...");
                            string expressionOwner = "'" + table.Name + "'[" + col.Name + "]";
                            string originalDaxExpression = col.Expression;
                            string formattedDaxExpression = DaxFormatter.FormatDaxExpression(originalDaxExpression, expressionOwner);
                            // write formatted expression back to calculated column
                            col.Expression = formattedDaxExpression;
                            // store hash of formatted expression for later comparison
                            string hashedDaxExpression = GetHashValueAsString(formattedDaxExpression);
                            if (col.Annotations.Contains("HashedExpression"))
                            {
                                col.Annotations["HashedExpression"].Value = hashedDaxExpression;
                            }
                            else
                            {
                                col.Annotations.Add(new Annotation { Name = "HashedExpression", Value = hashedDaxExpression });
                            }
                        }
                        else
                        {
                            Console.WriteLine(" - DAX already formatted");
                        }
                    }

                }

                // check which calculated tables require formatting
                if ((table.Partitions.Count > 0) &&
                    (table.Partitions[0].SourceType == PartitionSourceType.Calculated))
                {
                    Console.Write("Calculated table: " + table.Name);
                    if (CalculatedTableRequiresFormatting(table))
                    {
                        var source = table.Partitions[0].Source as CalculatedPartitionSource;
                        Console.WriteLine(" - formatting DAX...");
                        string expressionOwner = table.Name;
                        string originalDaxExpression = source.Expression;
                        string formattedDaxExpression = DaxFormatter.FormatDaxExpression(originalDaxExpression, expressionOwner);
                        // write formatted expression back to calculated column
                        source.Expression = formattedDaxExpression;
                        // store hash of formatted expression for later comparison
                        string hashedDaxExpression = GetHashValueAsString(formattedDaxExpression);
                        if (table.Annotations.Contains("HashedExpression"))
                        {
                            table.Annotations["HashedExpression"].Value = hashedDaxExpression;
                        }
                        else
                        {
                            table.Annotations.Add(new Annotation { Name = "HashedExpression", Value = hashedDaxExpression });
                        }

                    }
                    else
                    {
                        Console.WriteLine(" - DAX already formatted ");
                    }
                }

            }

            model.RequestRefresh(RefreshType.Automatic);
            model.SaveChanges();

            Console.WriteLine();
            Console.WriteLine("Press any key to continue");
            Console.ReadLine();
        }

        private static string FormatDaxExpression(string daxInput, string expressionOwner = "")
        {

            string prefix = string.IsNullOrEmpty(expressionOwner) ? "" : expressionOwner + " =";

            string daxExpression = prefix + daxInput;

            RequestBody requestBody = new RequestBody
            {
                CallerApp = "Apollo",
                Dax = daxExpression,
                DecimalSeparator = ".",
                ListSeparator = ",",
                DatabaseCompatibilityLevel = "520",
                SkipSpaceAfterFunctionName = false
            };

            string postBody = JsonConvert.SerializeObject(requestBody);

            HttpContent body = new StringContent(postBody);
            body.Headers.ContentType = new MediaTypeWithQualityHeaderValue("application/json");
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Add("Accept", "application/json; charset=UTF-8");
            HttpResponseMessage response = client.PostAsync(restUrl, body).Result;

            if (response.IsSuccessStatusCode)
            {
                string jsonResponse = response.Content.ReadAsStringAsync().Result;
                ResponseBody responseBody = JsonConvert.DeserializeObject<ResponseBody>(jsonResponse);
                string formattedExression = responseBody.formatted.Replace(prefix, "").Replace("\r\n", "\n");
                return formattedExression;
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine("OUCH! - error occurred during POST REST call");
                Console.WriteLine();
                return string.Empty;
            }
        }

        private static bool CalculatedColumnRequiresFormatting(CalculatedColumn column)
        {
            if (!column.Annotations.Contains("HashedExpression"))
            {
                return true;
            }
            string daxExpression = column.Expression;
            string hashedDaxExpression = GetHashValueAsString(daxExpression);
            string lastStoredHash = column.Annotations["HashedExpression"].Value;
            return hashedDaxExpression != lastStoredHash;
        }

        private static bool MeasureRequiresFormatting(Measure measure)
        {
            if (!measure.Annotations.Contains("HashedExpression"))
            {
                return true;
            }
            string daxExpression = measure.Expression;
            string hashedDaxExpression = GetHashValueAsString(daxExpression);
            string lastStoredHash = measure.Annotations["HashedExpression"].Value;
            return hashedDaxExpression != lastStoredHash;
        }

        private static bool CalculatedTableRequiresFormatting(Table table)
        {
            if (!table.Annotations.Contains("HashedExpression"))
            {
                return true;
            }
            var source = table.Partitions[0].Source as CalculatedPartitionSource;
            string daxExpression = source.Expression;
            string hashedDaxExpression = GetHashValueAsString(daxExpression);
            string lastStoredHash = table.Annotations["HashedExpression"].Value;
            return hashedDaxExpression != lastStoredHash;
        }

        private static string GetHashValueAsString(string input)
        {
            byte[] data = Encoding.UTF8.GetBytes(input);
            using (var shaM = new SHA1Managed())
            {
                byte[] result = shaM.ComputeHash(data);
                return Convert.ToBase64String(result);
            }
        }

    }

    class RequestBody
    {
        public string Dax { get; set; }
        public object MaxLineLenght { get; set; }
        public bool SkipSpaceAfterFunctionName { get; set; }
        public string ListSeparator { get; set; }
        public string DecimalSeparator { get; set; }
        public string CallerApp { get; set; }
        public string CallerVersion { get; set; }
        public string ServerName { get; set; }
        public string ServerEdition { get; set; }
        public string ServerType { get; set; }
        public string ServerMode { get; set; }
        public string ServerLocation { get; set; }
        public string ServerVersion { get; set; }
        public string DatabaseName { get; set; }
        public string DatabaseCompatibilityLevel { get; set; }
    }

    class ResponseBody
    {
        public string formatted { get; set; }
        public List<object> errors { get; set; }
    }


}
