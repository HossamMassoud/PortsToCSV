using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System.Reflection;

enum ColumnPosition
{
    Name = 1,
    InterfaceSelector,
    PortRange,
    AssociatedPolicyGroup
};
partial class Program
{
    static void Main(string[] args)
    {
        string ExcelFilesFolderPath = "";
        while (!Directory.Exists(ExcelFilesFolderPath))
        {
            if (!string.IsNullOrEmpty(ExcelFilesFolderPath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Cannot find the path ... ");
            }

            Console.ResetColor();
            Console.Write("Enter Excel files folder path (type Exit to exit): ");
            ExcelFilesFolderPath = Console.ReadLine();
            
            if (!ExcelFilesFolderPath.EndsWith("\\"))
            {
                ExcelFilesFolderPath += "\\";
            }

            if (ExcelFilesFolderPath.Trim().ToLower() == "exit")
            {
                Environment.ExitCode = -1;
                return;
            }
        }

        GenerateReports(ExcelFilesFolderPath);
    }

    private static void GenerateReports(string ExcelFilesFolderPath)
    {
        Application excel = new Application();

        try
        {
            var extension = new List<string> { "xlsx" };
            foreach (var excelFile in Directory.GetFiles(ExcelFilesFolderPath).Where(x => extension.Contains(Path.GetExtension(x).TrimStart('.').ToLowerInvariant())))
            {
                var fileName = Path.GetFileName(excelFile);
                Workbook workbook = excel.Workbooks.Open(excelFile);
                List<NetworkConfiguration> networkConfigurations = new List<NetworkConfiguration>();
                NetworkConfiguration networkConfiguration;
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    int TotalRowsWithHeader = worksheet.UsedRange.Rows.Count;
                    string networkConfigurationName = "";
                    string portRange;
                    for (int i = 2; i <= TotalRowsWithHeader; i++)
                    {
                        if (!string.IsNullOrEmpty((string)(worksheet.Cells[i, ColumnPosition.Name] as Microsoft.Office.Interop.Excel.Range).Value))
                        {
                            networkConfigurationName = (string)(worksheet.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value;
                        }
                        portRange = (string)(worksheet.Cells[i, ColumnPosition.PortRange] as Microsoft.Office.Interop.Excel.Range).Value;
                        int forwardSlashPosition = 0;
                        List<int> portList = new List<int>();
                        foreach (string Port in portRange.Split(',').Select(x => x.Trim()).ToList())
                        {
                            forwardSlashPosition = Port.IndexOf("/");
                            portList.Add(Convert.ToInt32(Port.Substring(forwardSlashPosition + 1)));
                        }
                        networkConfiguration = new NetworkConfiguration();
                        networkConfiguration.Name = networkConfigurationName;
                        networkConfiguration.InterfaceSelector = (string)(worksheet.Cells[i, ColumnPosition.InterfaceSelector] as Microsoft.Office.Interop.Excel.Range).Value;
                        networkConfiguration.AssociatedPolicyGroup = (string)(worksheet.Cells[i, ColumnPosition.AssociatedPolicyGroup] as Microsoft.Office.Interop.Excel.Range).Value;
                        networkConfiguration.PortFrom = portList.Min().ToString();
                        networkConfiguration.PortTo = portList.Max().ToString();
                        networkConfigurations.Add(networkConfiguration);
                    }
                    ExportData.ExportCsv(networkConfigurations, $"{ExcelFilesFolderPath}{fileName}.csv", true);
                }
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Report Generation Completed");
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Report Generation Failed ...");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Exception: ");
            Console.WriteLine(ex);
        }
        finally
        {
            excel.Quit();
        }

        Console.ResetColor();
        Console.WriteLine("Press enter to exit ...");
        Console.ReadLine();
    }
}

class NetworkConfiguration
{
    [JsonProperty("int-profile-name")]
    public String Name { get; set; }
    [JsonProperty("int-selector-name")]
    public String InterfaceSelector { get; set; }
    [JsonProperty("from-port-number")]
    public String PortFrom { get; set; }
    [JsonProperty("to-port-number")]
    public String PortTo { get; set; }
    [JsonProperty("int-policy-group")]
    public String AssociatedPolicyGroup { get; set; }
}

public static class ExportData
{
    public static void ExportCsv<T>(List<T> data, string fileName, Boolean UseJsonPropertAsHeader = false)
    {
        var content = new System.Text.StringBuilder();

        // Get Table Header
        var header = "";
        if (UseJsonPropertAsHeader)
        {
            header = string.Join(",",
                                 typeof(T).GetProperties()
                                          .Select(p => p.GetCustomAttribute<JsonPropertyAttribute>())
                                          .Select(jp => jp.PropertyName));
        }
        else
        {
            header = string.Join(",",
                                 typeof(T).GetProperties()
                                          .Select(p => p.Name));
        }
        content.AppendLine(header);

        // Get Table Data
        var properties = typeof(T).GetProperties();
        foreach (var row in data)
        {
            var line = "";
            foreach (var prop in properties)
            {
                line += prop.GetValue(row, null) + ",";
            }
            line = line.Substring(0, line.Length - 1);
            content.AppendLine(line);
        }

        // Save to File
        System.IO.File.WriteAllText(fileName, content.ToString(), Encoding.UTF8);
    }
}