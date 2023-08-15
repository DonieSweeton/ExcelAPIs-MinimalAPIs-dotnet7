using ExcelApi.Repositories.IRepository;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;

namespace ExcelApi.EndPoints
{
    public static class ExcelEndPoints
    {

        public static void ExcelApi(this WebApplication app)
        {
            app.MapGet("/api/export", async (IExcelRepository excelRepository) =>
            {
                try
                {
                    // Replace "fileName.xlsx" with the desired file name and extension.
                    string fileName = "myData.xlsx";

                    // Call the ExportExcel method from the IExcelRepository
                    excelRepository.ExportExcel(fileName);

                    // Return the Excel file as the response
                    var fileBytes = await File.ReadAllBytesAsync(fileName);
                    return Results.File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
                catch (Exception ex)
                {
                    // Handle the exception and return an error response.
                    return Results.BadRequest($"Error exporting data to Excel: {ex.Message}");
                }
            });


            app.MapPost("/api/import", async (IExcelRepository excelRepository, IFormFile file, IWebHostEnvironment hostingEnvironment) =>
            {
                try
                {
                    // Check if a file was uploaded.
                    if (file == null || file.Length == 0)
                        return Results.BadRequest("No file uploaded.");

                    // Validate that the uploaded file is an Excel file.
                    if (!file.ContentType.Equals("application/vnd.ms-excel") && !file.ContentType.Equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                        return Results.BadRequest("Invalid file format. Only Excel files are allowed.");

                    // Save the uploaded file to a temporary location on the server.
                    string uploadsFolder = Path.Combine(hostingEnvironment.ContentRootPath, "uploads");
                    Directory.CreateDirectory(uploadsFolder); // Create the folder if it doesn't exist
                    string uniqueFileName = $"{Guid.NewGuid()}-{file.FileName}";
                    string filePath = Path.Combine(uploadsFolder, uniqueFileName);

                    using (var fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        await file.CopyToAsync(fileStream);
                    }

                    // Call the ImportExcel method to process the uploaded file.
                    excelRepository.ImportExcel(filePath);

                    // Clean up the temporary file.
                    File.Delete(filePath);

                    // Return a success response.
                    return Results.Ok("Data imported from Excel successfully.");
                }
                catch (Exception ex)
                {
                    // Handle the exception and return an error response.
                    return Results.BadRequest($"Error importing data from Excel: {ex.Message}");

                }
            });
        }
    }
}
