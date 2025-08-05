using ClosedXML.Excel;
using ExcelDataReader;
using KhatiExcel.Model;
using Microsoft.AspNetCore.Http;
using System.Data;
using System.Text;

namespace KhatiExcel.Feature
{
    public class LoadExcel : ILoadExcel
    {
        public (bool success, List<ExcelModel>[]? data, string? message, string? errorMessage)
            Fetch(string path, bool header = true)
        {
            try
            {
                var file = path;
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = false,
                    }
                };

                DataSet dataSet = excelReader.AsDataSet(conf);
                DataTable dt = dataSet.Tables[0];
                var Wor = dataSet.Tables["Sheet1"];

                if (dataSet == null)
                {
                    return (false, null, "Error Occurred", "Excel Dataset fetch failed");
                }
                if (dataSet.Tables == null)
                {
                    return (false, null, "Error Occurred", "Excel Dataset fetch failed");
                }

                var noOfRowCount = dataSet?.Tables?["Sheet1"]?.Rows.Count;
                var colPosition = dataSet?.Tables?["Sheet1"]?.Columns.Count;

                List<string> ColumnHeader = new List<string>();

                if (header)
                {

                    for (int i = 0; i < 1; i++)
                    {
                        for (int j = 0; j < colPosition; j++)
                        {
                            var text = dataSet?.Tables[0].Rows[i][j].ToString();
                            if (text != null)
                                ColumnHeader.Add(text);
                        }
                    }
                }

                List<ExcelModel>[] listRows = new List<ExcelModel>
                    [Convert.ToInt32(header ? noOfRowCount - 1 : noOfRowCount)];

                for (int i = header ? 1 : 0; i < noOfRowCount; i++)
                {
                    List<ExcelModel> list = new List<ExcelModel>();
                    listRows[header ? i - 1 : i] = new List<ExcelModel>();

                    for (int j = 0; j < colPosition; j++)
                    {
                        var text = dataSet?.Tables[0].Rows[i][j];

                        var model = new ExcelModel()
                        {
                            ColumnName = header ? ColumnHeader[j] : "",
                            ColumnValue = text?.ToString(),
                            ColumnGroup = ((char)(65 + j)).ToString(),
                        };

                        listRows[header ? i - 1 : i].Add(model);
                    }
                }

                return (true, listRows, "Data Fetch Successfully", null);
            }

            catch (Exception ex)
            {
                return (false, null, "Error Occurred", ex.Message);
            }
        }

        public (bool success, List<ExcelModel>[]? data, string? message, string? errorMessage)
            Fetch(IFormFile file, bool header = true)
        {
            try
            {
                if (file == null || file.Length == 0)
                {
                    return (false, null, "Invalid file", "Uploaded file is empty or null");
                }

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                using var stream = file.OpenReadStream();
                using var excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = false,
                    }
                };

                DataSet dataSet = excelReader.AsDataSet(conf);
                if (dataSet == null || dataSet.Tables.Count == 0)
                {
                    return (false, null, "Error Occurred", "Excel Dataset fetch failed");
                }

                DataTable table = dataSet.Tables[0];
                int rowCount = table.Rows.Count;
                int colCount = table.Columns.Count;

                List<string> columnHeaders = new List<string>();

                if (header)
                {
                    for (int j = 0; j < colCount; j++)
                    {
                        var headerValue = table.Rows[0][j]?.ToString();
                        columnHeaders.Add(headerValue ?? $"Column{j + 1}");
                    }
                }

                int actualRowCount = header ? rowCount - 1 : rowCount;
                List<ExcelModel>[] listRows = new List<ExcelModel>[actualRowCount];

                for (int i = header ? 1 : 0; i < rowCount; i++)
                {
                    var rowList = new List<ExcelModel>();
                    for (int j = 0; j < colCount; j++)
                    {
                        var cellValue = table.Rows[i][j];
                        var model = new ExcelModel
                        {
                            ColumnName = header ? columnHeaders[j] : "",
                            ColumnValue = cellValue?.ToString(),
                            ColumnGroup = ((char)(65 + j)).ToString()
                        };
                        rowList.Add(model);
                    }

                    listRows[header ? i - 1 : i] = rowList;
                }

                return (true, listRows, "Data Fetch Successfully", null);
            }
            catch (Exception ex)
            {
                return (false, null, "Error Occurred", ex.Message);
            }
        }

        public (bool success, string? base64, string? message, string? errorMessage)
            ListToExcelBase64<T>(string SheetName, List<string> HeaderName, List<T> Rows)
            where T : class
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    IXLWorksheet worksheet =
                    workbook.Worksheets.Add(SheetName);

                    for (int i = 0; i < HeaderName.Count; i++)
                    {
                        worksheet.Cell(1, i + 1).Value = HeaderName[i];
                        worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                    }

                    for (int index = 1; index <= Rows.Count; index++)
                    {
                        int i = 1;
                        var t = Rows[index - 1];

                        foreach (var prop in t.GetType().GetProperties())
                        {
                            var s = prop.GetValue(t, null);
                            worksheet.Cell(index + 1, i).Value = prop.GetValue(t, null);
                            i++;
                        }
                    }

                    var stream = new MemoryStream();
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    string base64 = Convert.ToBase64String(content);
                    return (true, base64, "Base 64 File created from List", null);
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message, ex.ToString());
            }
        }
    }
}

