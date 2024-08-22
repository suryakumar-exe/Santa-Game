using OfficeOpenXml;
using SantaGame.Modal;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

public class Santa
{
    string EMP_FILE_PATH = @"C:\Users\KumarSu\source\repos\SantaGame\SantaGame\Employee-List.xlsx";
    public static async Task Main()
    {
        try
        {
            var employees = await FetchDatasFromExcel();
            if (employees == null || employees.Count == 0)
            {
                Console.WriteLine("No employee data found.");
                return;
            }
            var gameResult = await SecretGameConsole(employees);
            if (gameResult == null || gameResult.Count == 0)
            {
                Console.WriteLine("Secret Santa game could not be processed.");
                return;
            }
            ResultFile(employees, gameResult);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    public static async Task<List<Employees>> FetchDatasFromExcel()
    {
        string filePath = @"C:\Users\KumarSu\source\repos\SantaGame\SantaGame\Employee-List.xlsx";
        List<Employees> employeeList = new List<Employees>();

        if (!File.Exists(filePath))
        {
            Console.WriteLine("Excel file not found.");
            return employeeList;
        }
        try
        {
            using (var excelPackage = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // For non-commercial use
                var worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    Console.WriteLine("No worksheet found in the Excel file.");
                    return employeeList;
                }
                int rowCount = worksheet.Dimension.End.Row;
                for (int row = 2; row <= rowCount; row++)
                {
                    var name = worksheet.Cells[row, 1].Value?.ToString();
                    var email = worksheet.Cells[row, 2].Value?.ToString();
                    if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(email))
                    {
                        employeeList.Add(new Employees { Name = name, Email = email });
                    }
                    else
                    {
                        Console.WriteLine($"Invalid data at row {row}. Skipping this row.");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to read Excel file: {ex.Message}");
        }
        return employeeList;
    }
    public static async Task<Dictionary<Employees, Employees>> SecretGameConsole(List<Employees> employeeList)
    {
        Dictionary<Employees, Employees> santaResult = new Dictionary<Employees, Employees>();
        List<Employees> clonedList = new List<Employees>(employeeList);
        Random random = new Random();
        try
        {
            foreach (var employee in employeeList)
            {
                Employees randomEmployee;
                bool lastTime;
                do
                {
                    if (clonedList.Count == 0)
                    {
                        Console.WriteLine("No more employees available to assign.");
                        return santaResult;
                    }
                    randomEmployee = clonedList[random.Next(clonedList.Count)];
                    lastTime = await CheckforLasttime(employee.Email, randomEmployee.Email);
                }
                while (randomEmployee.Email == employee.Email && !lastTime && clonedList.Count > 1);

                if (!santaResult.ContainsKey(employee))
                {
                    santaResult.Add(employee, randomEmployee);
                    clonedList.Remove(randomEmployee);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during Secret Santa assignment: {ex.Message}");
        }

        return santaResult;
    }

    public static void ResultFile(List<Employees> employeeList, Dictionary<Employees, Employees> gameResult)
    {
        //Replace with your file path
        string filePath = @"C:\Users\KumarSu\source\repos\SantaGame\SantaGame\SecretSanta-Result.xlsx";

        try
        {
            using (var package = new ExcelPackage())
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // For non-commercial use
                var worksheet = package.Workbook.Worksheets.Add("Secret Santa Result");

                worksheet.Cells[1, 1].Value = "Employee_Name";
                worksheet.Cells[1, 2].Value = "Employee_EmailID";
                worksheet.Cells[1, 3].Value = "Secret_Child_Name";
                worksheet.Cells[1, 4].Value = "Secret_Child_EmailID";

                int row = 2;

                foreach (var pair in gameResult)
                {
                    worksheet.Cells[row, 1].Value = pair.Key.Name;
                    worksheet.Cells[row, 2].Value = pair.Key.Email;
                    worksheet.Cells[row, 3].Value = pair.Value.Name;
                    worksheet.Cells[row, 4].Value = pair.Value.Email;
                    row++;
                }

                FileInfo file = new FileInfo(filePath);
                package.SaveAs(file);
            }

            Console.WriteLine($"Secret Santa results saved to {filePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save the result file: {ex.Message}");
        }
    }

    public static async Task<bool> CheckforLasttime(string emp, string randomEmail)
    {
        // Perform the activity of previous year match
        // Check if the emp matches with random email and return true, otherwise false
        return false;
    }
}
