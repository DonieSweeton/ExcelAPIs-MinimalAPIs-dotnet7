namespace ExcelApi.Repositories.IRepository
{
    public interface IExcelRepository
    {
        void ExportExcel(string fileName);
        void ImportExcel(string fileName);
    }
}
