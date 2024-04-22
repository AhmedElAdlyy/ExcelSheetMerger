using ExcelSheetMerger.ViewModels;
using OfficeOpenXml;

namespace ExcelSheetMerger
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Excel Sheet has headers? \n 1- ignore headers \n 2- keep headers");
            var hasHeader = Console.ReadLine();

            bool ignoreHeaders = hasHeader == "1" ? true : false;


            PathsViewModel Paths = new PathsViewModel
            {
                Source = "Files\\Source\\",
                Destination = "Files\\Destination\\"
            };

            Paths = GetFilesName(Paths);

            var times = new TimeViewModel
            {
                StartTime = DateTime.Now,
            };

            List<ResultSummary> results = ReadData(Paths, ignoreHeaders);

            times.EndTime = DateTime.Now;
            Console.WriteLine($"Processing Done Successfully in {times.EndTime - times.StartTime}");

            Console.WriteLine("Creating Final Results...");

            times.StartTime = DateTime.Now;


            List<ResultSummary> finalResults = FinalizeResult(results, Paths.Destination, ignoreHeaders);


            times.EndTime = DateTime.Now;


            Console.WriteLine($"Operation Done Successfully in {times.EndTime - times.StartTime}");

            foreach (ResultSummary result in finalResults)
            {
                Console.WriteLine($"{result.Count} rows {result.Status}");
            }

            Console.ReadLine();
        }

        private static PathsViewModel GetFilesName(PathsViewModel folders)
        {
            if (Directory.Exists(folders.Source))
            {
                string fileName = Directory.GetFiles(folders.Source).First();
                folders.Source = fileName;
            }

            if (Directory.Exists(folders.Destination))
            {
                string fileName = Directory.GetFiles(folders.Destination).First();
                folders.Destination = fileName;
            }

            return folders;
        }

        private static List<ResultSummary> ReadData(PathsViewModel paths, bool ignoreHeaders)
        {
            List<ResultSummary> results = new List<ResultSummary>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage source = new ExcelPackage(new FileInfo(paths.Source)))
            {
                ExcelWorksheet worksheet = source.Workbook.Worksheets[0];

                int countRow = worksheet.Dimension.Rows;
                int countCol = worksheet.Dimension.Columns;

                Dictionary<string, string?> sourceData = new Dictionary<string, string?>();

                int row = 0;

                for (row = ignoreHeaders == true ? 2 : 1; row <= countRow; row++)
                {

                    string ID = worksheet.Cells[row, 1].Value?.ToString();
                    string Name = worksheet.Cells[row, 2].Value?.ToString();

                    sourceData.Add(ID, Name);

                    var result = UpdateDestination(paths.Destination, ignoreHeaders, sourceData);
                    ResultSummary summary = new ResultSummary
                    {
                        ReturnValue = result
                    };

                    results.Add(summary);
                }
            }

            return results;
        }

        private static bool UpdateDestination(string path, bool ignoreHeaders, Dictionary<string, string> sourceData)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage dest = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet worksheet = dest.Workbook.Worksheets[0];
                int countRow = worksheet.Dimension.Rows;
                int countCol = worksheet.Dimension.Columns;

                int row = 0;

                for (row = ignoreHeaders == true ? 2 : 1; row <= countRow; row++)
                {


                    string id = worksheet.Cells[row, 1].Value?.ToString();

                    if (sourceData.ContainsKey(id))
                    {
                        worksheet.Cells[row, 2].Value = sourceData[id];
                        dest.Save();
                        return true;
                    }


                }

                return false;
            }
        }

        private static List<ResultSummary> FinalizeResult(List<ResultSummary> results, string path, bool ignoreHeaders)
        {
            int Updated = results.Where(w => w.ReturnValue == true).Count();
            int NotFoundInDestination = results.Where(w => w.ReturnValue == false).Count();
            int NotInSource = CountEmptyIds(path, ignoreHeaders);

            results.Clear();
            results.Add(new ResultSummary
            {
                Count = Updated,
                Status = "Updated Successfully"
            });
            results.Add(new ResultSummary
            {
                Count = NotFoundInDestination,
                Status = "Not Exist In The Copy"
            });
            results.Add(new ResultSummary
            {
                Count = NotInSource,
                Status = "Not Exist In The Source"
            });

            return results;
        }

        private static int CountEmptyIds(string path, bool ignoreHeaders)
        {
            int count = 0;

            using (ExcelPackage dest = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet worksheet = dest.Workbook.Worksheets[0];
                int countRow = worksheet.Dimension.Rows;
                int countCol = worksheet.Dimension.Columns;

                int row = 0;

                for (row = ignoreHeaders == true ? 2 : 1; row <= countRow; row++)
                {
                    string cellValue = worksheet.Cells[row, 2].Value?.ToString();

                    if (cellValue == null)
                        count++;
                }

                return count;
            }
        }


    }
}
