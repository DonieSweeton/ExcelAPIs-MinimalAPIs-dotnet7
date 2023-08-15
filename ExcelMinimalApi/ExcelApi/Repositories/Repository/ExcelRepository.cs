using ExcelApi.Data;
using ExcelApi.Models;
using ExcelApi.Repositories.IRepository;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelApi.Repositories.Repository
{
    public class ExcelRepository : IExcelRepository
    {
        private readonly AppDbContext _dbContext; // The database context that is used to access the data.

        public ExcelRepository(AppDbContext dbContext)
        {
            _dbContext = dbContext;
        }

        public void ExportExcel(string fileName)
        {
            // Create a new Excel package.
            using var package = new ExcelPackage();

            // Get the list of distinct group IDs from the groups table.
            List<int?> groupIds = _dbContext.Users.Where(u => u.IsDeleted == false)
                                                 .Where(u => u.GroupId != null)
                                                .Select(u => u.GroupId)
                                                .Distinct()
                                                .ToList();

            foreach (int? groupId in groupIds)
            {
                // Get the group details from the groups table.
                Group? group = _dbContext.Groups.FirstOrDefault(g => g.GroupId == groupId);

                if (group == null) 
                {
                    continue;
                }

                // Create a new sheet for each group.
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add(group.GroupName);

                // Add the additional details for the group.
                sheet.Cells["A1"].Value = "Group Name:";
                sheet.Cells["B1"].Value = group.GroupName;

                sheet.Cells["A2"].Value = "Group Owner:";
                sheet.Cells["B2"].Value = group.CreatedBy;

                sheet.Cells["A3"].Value = "Date:";
                sheet.Cells["B3"].Value = DateTime.Now.ToString("yyyy-MM-dd");

                // Add the title row with borders for each sheet.
                var titleRow = sheet.Cells[5, 1, 5, 6];
                titleRow.Style.Font.Bold = true;
                titleRow.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                titleRow.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                titleRow.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                titleRow.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                titleRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                titleRow.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

                sheet.Cells[5, 1].Value = "User Name";
                sheet.Cells[5, 2].Value = "User Email";
                sheet.Cells[5, 3].Value = "Created By";
                sheet.Cells[5, 4].Value = "Created Date";
                sheet.Cells[5, 5].Value = "Modified By";
                sheet.Cells[5, 6].Value = "Modified Date";

                // Get the list of users for the current group.
                List<User> users = _dbContext.Users.Where(u => u.GroupId == groupId && u.IsDeleted == false).ToList();

                // Iterate over the list of users and add them to the sheet with borders.
                int row = 6;
                foreach (User user in users)
                {
                    var dataRow = sheet.Cells[row, 1, row, 6];
                    dataRow.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    dataRow.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    dataRow.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    dataRow.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    // Add the user's name to the sheet.
                    sheet.Cells[row, 1].Value = user.UserName;

                    // Add the user's email to the sheet.
                    sheet.Cells[row, 2].Value = user.UserEmail;

                    // Add the user's createdBy to the sheet.
                    sheet.Cells[row, 3].Value = user.CreatedBy;

                    // Add the user's createdDate to the sheet.
                    sheet.Cells[row, 4].Value = user.CreatedDate;

                    // Add the user's modifiedBy to the sheet.
                    if (user.ModifiedBy != null)
                    {
                        sheet.Cells[row, 5].Value = user.ModifiedBy;
                    }

                    // Add the user's modifiedDate to the sheet.
                    if (user.ModifiedDate != null)
                    {
                        sheet.Cells[row, 6].Value = user.ModifiedDate;
                    }

                    row++;
                }
            }

            // Save the package to a file.
            package.SaveAs(new FileInfo(fileName));
        }

        public void ImportExcel(string fileName)
        {
            
            // Load the Excel file into an ExcelPackage.
            using var package = new ExcelPackage(new FileInfo(fileName));

            // Get the worksheet from the package.
            ExcelWorksheet sheet = package.Workbook.Worksheets[0];

            // Get the last row in the worksheet that contains data.
            int lastRow = sheet.Dimension.End.Row;

            // Iterate through each row in the worksheet and insert/update the data into the database.
            for (int row = 2; row <= lastRow; row++)
            {
                // Read the data from the Excel cells.
                string userName = sheet.Cells[row, 1].GetValue<string>();
                string userEmail = sheet.Cells[row, 2].GetValue<string>();
                string createdBy = sheet.Cells[row, 3].GetValue<string>();
                int groupId = sheet.Cells[row, 4].GetValue<int>();

                // Validate that all columns and rows are filled and not null before processing the data.
                if (string.IsNullOrEmpty(userName) || string.IsNullOrEmpty(userEmail) || string.IsNullOrEmpty(createdBy) || groupId == 0)
                {
                    // If any column or row is null or empty, skip processing and continue to the next row.
                    continue;
                }

                // Check if a user with the given email already exists in the database.
                User? existingUser = _dbContext.Users.FirstOrDefault(u => u.UserEmail == userEmail);

                if (existingUser != null)
                {
                    // Update the existing user's properties.
                    existingUser.UserName = userName;
                    existingUser.CreatedBy = createdBy;
                    existingUser.CreatedDate = DateTime.Now;
                    existingUser.GroupId = groupId;
                    _dbContext.Users.Update(existingUser);
                }
                else
                {
                    // Create a new User object with the read data.
                    User newUser = new()
                    {
                        UserName = userName,
                        UserEmail = userEmail,
                        CreatedBy = createdBy,
                        CreatedDate = DateTime.Now,
                        GroupId = groupId
                    };

                    // Add the new User object to the database.
                    _dbContext.Users.Add(newUser);
                }
            }

            // Save the changes to the database.
            _dbContext.SaveChanges();



        }
    }
}
