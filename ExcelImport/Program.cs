using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using GemBox.Spreadsheet;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace ExcelImport
{
    public class Program
    {
        public static void Main(string[] args)
        {
            CreateHostBuilder(args).Build().Run();
        }

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .ConfigureWebHostDefaults(webBuilder =>
                {
                    webBuilder.UseStartup<Startup>();
                });
    }
}

//public class Program
//{
//    public static void Main(string[] args)
//    {
//        CreateHostBuilder(args).Build().Run();
//    }

//    public static IHostBuilder CreateHostBuilder(string[] args) =>
//        Host.CreateDefaultBuilder(args)
//            .ConfigureWebHostDefaults(webBuilder =>
//            {
//                webBuilder.UseStartup<Startup>();
//            });
//}


//class Program
//{
//    static void Main()
//    {
//        // If using Professional version, put your serial key below.
//        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

//        var workbook = ExcelFile.Load("SimpleTemplate.xlsx");

//        var worksheet = workbook.Worksheets[0];

//        int columnCount = worksheet.CalculateMaxUsedColumns();
//        for (int i = 0; i < columnCount; i++)
//            worksheet.Columns[i].AutoFit(1, worksheet.Rows[1], worksheet.Rows[worksheet.Rows.Count - 1]);

//        workbook.Save("Row_Column AutoFit.xlsx");
//    }
//}