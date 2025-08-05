using KhatiExcel.Model;
using Microsoft.AspNetCore.Http;

namespace KhatiExcel.Feature
{
    public interface ILoadExcel
    {
        (bool success, List<ExcelModel>[]? data, string? message, string? errorMessage)
            Fetch(string path, bool header = true);

        (bool success, List<ExcelModel>[]? data, string? message, string? errorMessage)
            Fetch(IFormFile file, bool header = true);

        (bool success, string? base64, string? message, string? errorMessage)
            ListToExcelBase64<T>(string SheetName, List<string> HeaderName, List<T> Rows) where T : class;
    }
}
