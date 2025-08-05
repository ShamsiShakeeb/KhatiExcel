using KhatiExcel.Model;
using Microsoft.AspNetCore.Http;

namespace KhatiExcel.Feature
{
    public interface ILoadExcel
    {
        (bool success, List<ExcelModel>[]? data, string? message, string? errorMessage)
            Fetch(string path, string sheetName, bool header = true);

        (bool success, List<ExcelModel>[]? data, string? message, string? errorMessage)
            Fetch(IFormFile file, string sheetName, bool header = true);

        (bool success, string? base64, string? message, string? errorMessage)
            ListToExcelBase64<T>(string SheetName, List<string> HeaderName, List<T> Rows) where T : class;
    }
}
