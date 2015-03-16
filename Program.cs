using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;

namespace SQL_postcode_updater
{
    class Program
    {
        private const string Primary = "Primary";
        private const string Home = "Home";

        static void Main(string[] args)
        {
            var book = new LinqToExcel.ExcelQueryFactory(@"Postcodes.xlsx");

            var query = book.Worksheet("Sheet1").Select(row => new
            {
                UserName = row["Username"].Cast<string>(),
                Name = row["Name"].Cast<string>(),
                PostcodeType = row["Postcode type"].Cast<string>(),
                Postcode = row["Postcode"].Cast<string>()
            });

            Console.WriteLine("Data received");

            using (var fileWriter = new StreamWriter("Script.txt"))
            {
                fileWriter.WriteLine("DELETE FROM ruth.AreaPostcodes");
                fileWriter.WriteLine();

                foreach (var data in query)
                {
                    if (string.Equals(data.PostcodeType, Home))
                    {
                        continue;
                    }

                    var areaId = string.Equals(data.PostcodeType, Primary) ? "PrimaryAreaId" : "SecondaryAreaId";

                    fileWriter.WriteLine("Go");
                    fileWriter.WriteLine("Insert Into ruth.AreaPostcodes(AreaId, PostcodeId)");
                    fileWriter.WriteLine("Values(");
                    fileWriter.WriteLine("(select {0} from ruth.Officers where UserName like '{1}'), ", areaId, data.UserName);
                    fileWriter.WriteLine("(select Id from ruth.Postcodes where Code like '{0}'));", data.Postcode);
                    fileWriter.WriteLine("Go");
                    fileWriter.WriteLine();
                }   
            }

            Console.WriteLine("Data saved");
            Console.ReadLine();
        }
    }
}
