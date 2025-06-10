using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
public class ExportFastExcelCS
{
    private static Dictionary<string, string> html_conv_char = new Dictionary<string, string>()
    {
        { "&","&amp;" },
        { "<","&lt;" },
        { ">","&gt;" },
        { "\r\n","<br>" },
        { "\n","<br>" },
        { "\r","<br>" }
    };
    // Helper function to escape special characters for XML
    private string ReplaceSpecialCharacter(string str)
    {
        return Regex.Replace(str, "[&<>]", new MatchEvaluator((Match match) => html_conv_char[match.Value]));
    }
    // Helper function to convert column index to Excel column reference (A, B, AA, etc.)
    private string ConvertColumnIndexToRef(int ColumnIndex)
    {
        string strRef = "";
        int col = ColumnIndex;

        while (true)
        {
            int x = col % 26;
            strRef = (char)(65 + x) + strRef;
            col = (int)(col / 26) - 1;
            if (col < 0)
            {
                break;
            }
        }
        columnRefCache[col] = strRef;
        return strRef;
    }
    private Dictionary<int, string> columnRefCache = new Dictionary<int, string>();
    // Helper function to get cell reference (e.g., A1, B2)
    private string GetCellRef(int Row, int Col)
    {
        if (!columnRefCache.ContainsKey(Col))
        {
            columnRefCache[Col] = ConvertColumnIndexToRef(Col);
        }
        return string.Format("{0}{1}", columnRefCache[Col], Row + 1);
    }
    private static DateTime excelStartTime = new DateTime(1899, 12, 30);
    // # Function to build cell elements for a row
    private string BuildCellElements(int rowIndex, object[] Values)
    {
        string ret = "";
        for (int colIndex = 0; colIndex < Values.Length; colIndex++)
        {
            string cellRef = GetCellRef(rowIndex, colIndex);
            object value = Values[colIndex];
            string typeName = value.GetType().FullName;
            if (typeName.Contains("Int")
             || typeName == "System.Decimal"
             || typeName == "System.Double"
             || typeName == "System.Single"
             )
            {
                ret += string.Format("<c r=\"{0}\" s=\"0\" t=\"n\"><v>{1}</v></c>", cellRef, value);
            }
            else if (typeName == "System.Boolean")
            {
                ret += string.Format("<c r=\"{0}\" s=\"0\" t=\"n\"><v>{1}</v></c>", cellRef, (bool)value ? 0 : 1);
            }
            else if (typeName == "System.DateTime")
            {
                ret += string.Format("<c r=\"{0}\" s=\"1\"><v>{1}</v></c>", cellRef, ((DateTime)value - excelStartTime).TotalSeconds / 86400);
            }
            else if (typeName == "System.String")
            {
                ret += string.Format("<c r=\"{0}\" s=\"0\" t=\"str\"><v>{1}</v></c>", cellRef, ReplaceSpecialCharacter((string)value));
            }
            else
            {
                ret += string.Format("<c r=\"{0}\" s=\"0\" t=\"str\"><v>{1}</v></c>", cellRef, ReplaceSpecialCharacter(value.ToString()));
            }
        }
        // Excel row index starts from 1, so add 1 to $RowIndex
        return string.Format("<row r=\"{0}\" spans=\"1:{1}\" customFormat=\"false\">{2}</row>", rowIndex + 1, Values.Length, ret);
    }

    private void WriteCellElements(System.IO.Stream stream, int rowIndex, object[] Values)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(BuildCellElements(rowIndex, Values));
        stream.Write(bytes, 0, bytes.Length);
    }
    private List<string> tableHeaders = new List<string>();
    private int startRowIndex = 0;
    // Determine headers
    private void DetermineHeaders(object[][] data, out List<string> headers)
    {
        headers = new List<string>();
        startRowIndex = 0;
        object[] firstRowData = data[0];
        if (firstRowData[0] is System.Data.DataRow)
        {
            System.Data.DataRow dataRow = (System.Data.DataRow)firstRowData[0];
            foreach (System.Data.DataColumn column in dataRow.Table.Columns)
            {
                headers.Add(column.ColumnName);
            }
        }
        else if (firstRowData[0] is System.Management.Automation.PSObject)
        {
            System.Management.Automation.PSObject psObj = (System.Management.Automation.PSObject)firstRowData[0];
            foreach (var prop in psObj.Properties)
            {
                headers.Add(prop.Name);
            }
        }
        else
        {
            foreach (object item in firstRowData)
            {
                headers.Add(item.ToString());
            }
            startRowIndex = 1;
        }
    }
    // Generate XML parts for headers
    private Dictionary<string, string> GenerateHeaders(object[][] data)
    {
        List<string> headers;
        DetermineHeaders(data, out headers);
        int dataCount = data.Length;
        Dictionary<string, int> used = new Dictionary<string, int>();
        string headerMap = "";
        string tableColumn = "";
        int index = 0;
        tableHeaders.Clear();
        foreach (string header in headers)
        {
            string name = header;
            while (tableHeaders.Contains(name))
            {
                if (!used.ContainsKey(header))
                {
                    used[header] = 0;
                }
                used[header]++;
                name = string.Format("{0} ({1})", header, used[header]);
            }
            index++;
            tableHeaders.Add(name);
            string xml_str = name.Replace("\"", "&quot;");
            tableColumn += string.Format("<tableColumn id=\"{0}\" name=\"{1}\"/>", index, xml_str);
            headerMap += string.Format("<si><t>{0}</t></si>", ReplaceSpecialCharacter(name)); // Already escaped in $headers
        }
        string rangeReference = GetCellRef(dataCount - startRowIndex, headers.Count - 1);
        return new Dictionary<string, string>()
        {
            { "header.size" , headers.Count.ToString() },
            { "header.map", headerMap },
            { "rangeReference", rangeReference },
            { "tableColumn", tableColumn },
            { "dimension", string.Format("<dimension ref=\"A1:{0}\"/>", rangeReference) },
        };
    }
    // Write Sheet
    private void WriteSheet(System.IO.Stream stream, object[][] data)
    {
        WriteCellElements(stream, 0, tableHeaders.ToArray());
        int rowIndex = 0;
        for (int row = startRowIndex; row < data.Length; row++)
        {
            object[] rowData = data[row];
            rowIndex++;
            // If the input was PSObject, we need to extract values in header order
            if (rowData[0] is System.Data.DataRow)
            {
                List<object> orderedValues = new List<object>();
                System.Data.DataRow dataRow = (System.Data.DataRow)rowData[0];
                foreach (string headerProp in tableHeaders)
                {
                    orderedValues.Add(dataRow[headerProp]);
                }
                WriteCellElements(stream, rowIndex, orderedValues.ToArray());
            }
            else if (rowData[0] is System.Management.Automation.PSObject)
            {
                List<object> orderedValues = new List<object>();
                System.Management.Automation.PSObject psObj = (System.Management.Automation.PSObject)rowData[0];
                foreach (string headerProp in tableHeaders)
                {
                    var prop = psObj.Properties[headerProp];
                    orderedValues.Add(prop != null ? prop.Value : null);
                }
                WriteCellElements(stream, rowIndex, orderedValues.ToArray());
            }
            else
            {
                WriteCellElements(stream, rowIndex, rowData);
            }
        }
        string closeContent = "</sheetData><tableParts count=\"1\"><tablePart r:id=\"rId1\"/></tableParts></worksheet>";
        byte[] bytes = Encoding.UTF8.GetBytes(closeContent);
        stream.Write(bytes, 0, bytes.Length);
    }
    // Create Zip file and add content
    private void CreateZipFileAndAddContent(string rootFolder, string filename, object[][] data)
    {
        System.IO.Compression.ZipArchive zipArchive = null;
        try
        {
            Dictionary<string, string> replacementRules = GenerateHeaders(data);
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }
            zipArchive = System.IO.Compression.ZipFile.Open(filename, System.IO.Compression.ZipArchiveMode.Create);
            string sourceFolderFullPath = System.IO.Path.GetFullPath(rootFolder);
            string[] files = System.IO.Directory.GetFiles(sourceFolderFullPath, "*", System.IO.SearchOption.AllDirectories);
            foreach (var file in files)
            {
                string fileContent = File.ReadAllText(file, Encoding.UTF8);
                fileContent = Regex.Replace(fileContent, "#{@([^}]+)}",
                 new MatchEvaluator((Match match) => replacementRules.ContainsKey(match.Groups[1].Value) ? replacementRules[match.Groups[1].Value] : ""));
                string relativePath = file.Replace("\\", "/").Substring(sourceFolderFullPath.Length).TrimStart('/');
                // ファイルを追加
                var entry = zipArchive.CreateEntry(relativePath, System.IO.Compression.CompressionLevel.Optimal); // SmallestSize)
                var entryStream = entry.Open();
                var bytes = Encoding.UTF8.GetBytes(fileContent);
                entryStream.Write(bytes, 0, bytes.Length);
                if (file.EndsWith("sheet1.xml"))
                {
                    WriteSheet(entryStream, data);
                }
                entryStream.Close();
            }
        }
        catch (System.Exception e)
        {
            Console.WriteLine(e);
        }
        finally
        {
            if (null != zipArchive)
            {
                zipArchive.Dispose();
            }
        }
    }
    // Export Excel File
    public static void Export(string rootFolder, string filename, object[][] data)
    {
        ExportFastExcelCS exportFastExcelCS = new ExportFastExcelCS();
        exportFastExcelCS.CreateZipFileAndAddContent(rootFolder, filename, data);
    }
}