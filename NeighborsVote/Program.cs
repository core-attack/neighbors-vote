using OfficeOpenXml;
using System.Reflection;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace NeighborsVote
{
    internal class Program
    {
        private static readonly object LockObject = new object();

        private const string InputFolder = "input";
        private const string TemplateSubPath = "template.docx";
        private const string ExcelFilePath = "весна 2 реест.xlsx";
        private const string NameAnchor = "__________________________________________________________________________________";

        private const int FirstRowWithData = 2;
        private const int NameColumnIndex = 5;
        private const int IdColumnIndex = 1;
        private const int FlatColumnIndex = 1;
        private const int AreaColumnIndex = 3;
        private const int ShareColumnIndex = 6;
        private const int BasisColumnIndex = 4;
        private static int BatchSize = 50;
        private static string SplitNamesSeparator = "/";
        private static string SkipMarker = "данные о правообладателе отсутствуют";

        private static string OutputSubPath(string x) => $"output_{x}.docx";

        public static async Task Main()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var assemblyLocation = Assembly.GetExecutingAssembly().Location;
            var directory = Path.GetDirectoryName(assemblyLocation);
            var projectDirectory = Directory.GetParent(directory).Parent.Parent;
            var root = Path.Combine(projectDirectory.FullName, InputFolder);

            if (!Directory.Exists(root))
            {
                Directory.CreateDirectory(root);
            }

            var templatePath = Path.Combine(root, TemplateSubPath);
            var excelFilePath = Path.Combine(root, ExcelFilePath);

            if (!Path.Exists(templatePath))
            {
                Console.Error.WriteLine($"Template file is missing here {root}");
                return;
            }

            if (!Path.Exists(excelFilePath))
            {
                Console.Error.WriteLine($"Data file is missing here {root}");
                return;
            }

            var customValues = ReadValuesFromExcel(excelFilePath);

            await CopyTemplateWithCustomValues(templatePath, root, customValues);
        }

        static IGrouping<string, Neighbor>[] ReadValuesFromExcel(string filePath)
        {
            var values = new List<Neighbor>();

            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[0];

            for (int row = FirstRowWithData; row <= worksheet.Dimension.End.Row; row++)
            {
                // Read value from the first column (Column A)
                var id = worksheet.Cells[row, IdColumnIndex].Text;
                var name = worksheet.Cells[row, NameColumnIndex].Text;
                var flat = worksheet.Cells[row, FlatColumnIndex].Text;
                double.TryParse(worksheet.Cells[row, AreaColumnIndex].Text, out var area);
                double.TryParse(worksheet.Cells[row, ShareColumnIndex].Text, out var share);
                var basis = worksheet.Cells[row, BasisColumnIndex].Text;

                if (name.Trim().Equals(SkipMarker))
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(id))
                {
                    if (name.Contains(SplitNamesSeparator))
                    {
                        var splitedNames = name.Split(SplitNamesSeparator);
                        var sharePart = share / splitedNames.Length;

                        foreach (var n in splitedNames)
                        {
                            var trim = n.Trim();

                            if (!string.IsNullOrWhiteSpace(trim))
                            {
                                values.Add(new Neighbor(Name: trim, Id: id, FlatNumber: flat, Area: area, Share: sharePart, Basis: basis));
                            }
                        }
                    }
                    else
                    {
                        values.Add(new Neighbor(Name: name, Id: id, FlatNumber: flat, Area: area, Share: share, Basis: basis));
                    }
                }
            }

            return values.GroupBy(x => x.Name).ToArray();
        }

        public static async Task CopyTemplateWithCustomValues(string templatePath, string root, IGrouping<string, Neighbor>[] customValues)
        {
            if (customValues.Length == 0)
            {
                Console.WriteLine("No custom values provided.");
                return;
            }

            using var originalDocument = DocX.Load(templatePath);
            var tables = originalDocument.Tables;

            if (tables.Count == 0)
            {
                Console.WriteLine("No tables found in the template document.");
                return;
            }

            using var templateClone = originalDocument.Copy();

            var skip = 0;
            var processedRecords = 0;
            var timestamp = DateTime.UtcNow.Ticks;

            do
            {
                Console.WriteLine($"Analysing batch from {skip} to {skip + BatchSize}");
                var batch = customValues.Skip(skip).Take(BatchSize);
                var outputPath = Path.Combine(root, OutputSubPath($"{skip}_{skip + BatchSize}_{timestamp}"));

                using var newDocument = DocX.Create(outputPath);

                var options = new ParallelOptions
                {
                    MaxDegreeOfParallelism = Environment.ProcessorCount * 2
                };

                await Task.WhenAll(Parallel.ForEachAsync(batch, options, async (neighborGroup, cancellationToken) =>
                {
                    Console.WriteLine($"Working with {neighborGroup.Key}");

                    lock (LockObject)
                    {
                        var clonedDocument = templateClone.Copy();

                        var nameParagraph = clonedDocument.Paragraphs.FirstOrDefault(x => x.Text.Equals(NameAnchor));

                        if (nameParagraph != null)
                        {
                            nameParagraph.ReplaceText(new StringReplaceTextOptions()
                                { SearchValue = NameAnchor, NewValue = neighborGroup.Key });
                        }

                        var clonedTable = clonedDocument.Tables[0];
                        var totalItems = neighborGroup.Count();
                        var emptyRowIndex = 2;

                        var repeats = totalItems - 1;

                        while (repeats > 0)
                        {
                            var prev = clonedTable.Rows[emptyRowIndex];
                            clonedTable.InsertRow(prev, true);
                            repeats--;
                        }

                        foreach (var neighbor in neighborGroup)
                        {
                            clonedTable.Rows[emptyRowIndex].Cells[0].Paragraphs[0].InsertText(neighbor.FlatNumber);
                            clonedTable.Rows[emptyRowIndex].Cells[1].Paragraphs[0].InsertText(neighbor.Area.ToString());
                            clonedTable.Rows[emptyRowIndex].Cells[2].Paragraphs[0].InsertText(neighbor.Share.ToString());
                            clonedTable.Rows[emptyRowIndex].Cells[3].Paragraphs[0].InsertText(neighbor.Basis);
                            clonedTable.Rows[emptyRowIndex].Cells[4].Paragraphs[0].InsertText(string.Empty);
                            emptyRowIndex++;
                        }

                        newDocument.InsertDocument(clonedDocument);
                    }

                    Console.WriteLine($"Completed with {neighborGroup.Key}");
                    Interlocked.Increment(ref processedRecords);
                }));

                newDocument.Sections[0].Remove();
                newDocument.Save();
                Console.WriteLine($"Document created successfully at {outputPath}");
                skip += BatchSize;
            } 
            while (processedRecords < customValues.Length);
        }

        public record Neighbor(string Name, string Id, string FlatNumber, double Area, double Share, string Basis);
    }
}
