using Microsoft.ML;
using Microsoft.ML.Data;
using OfficeOpenXml;
using Microsoft.Extensions.Configuration;

namespace TPDMAutomationProvider
{
    // Data Models
    public class CommentData
    {
        [LoadColumn(0)]
        public string Comment { get; set; } = string.Empty;

        [LoadColumn(1)]
        public string Action { get; set; } = string.Empty;
    }

    public class CommentPrediction
    {
        [ColumnName("PredictedLabel")]
        public string PredictedAction { get; set; } = string.Empty;
    }

    public class ExcelRowData
    {
        public Dictionary<string, object> ColumnValues { get; set; } = new Dictionary<string, object>();
        public string DelegateComment { get; set; } = string.Empty;
        public string PredictedAction { get; set; } = string.Empty;
        public string SourceSheet { get; set; } = string.Empty;
        public int RowNumber { get; set; }
    }

    /// <summary>
    /// TPDM Automation Provider - ML.NET based Excel processor
    /// Processes Excel files to predict actions based on delegate comments
    /// </summary>
    class Program
    {
        private static MLContext _mlContext = new MLContext(seed: 0);
        private static ITransformer? _trainedModel;

        static async Task<int> Main(string[] args)
        {
            try
            {
                Console.WriteLine("========================================");
                Console.WriteLine("   TPDM Automation Provider v1.0");
                Console.WriteLine("========================================");
                Console.WriteLine();

                // Configuration
                var inputPath = "./TestData/input.xlsx";
                var outputPath = "./Output/";
                var modelPath = "./MLModels/comment_classifier.zip";

                // Ensure directories exist
                Directory.CreateDirectory("./TestData");
                Directory.CreateDirectory("./Output");
                Directory.CreateDirectory("./MLModels");

                // Create sample input file if it doesn't exist
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine("Creating sample input file...");
                    await CreateSampleInputFile(inputPath);
                    Console.WriteLine($"Sample file created: {inputPath}");
                }

                // Train or load ML model
                if (!File.Exists(modelPath))
                {
                    Console.WriteLine("Training ML model...");
                    await TrainModel(modelPath);
                }
                else
                {
                    Console.WriteLine("Loading existing ML model...");
                    LoadModel(modelPath);
                }

                // Process Excel file
                Console.WriteLine("Processing Excel file...");
                var excelData = await ReadExcelFile(inputPath);
                
                if (!excelData.Any())
                {
                    Console.WriteLine("No data found in Excel file.");
                    return 1;
                }

                // Make predictions
                Console.WriteLine("Making predictions...");
                var processedData = ProcessData(excelData);

                // Display summary
                DisplaySummary(processedData);

                // Create output files
                Console.WriteLine("Creating output files...");
                await CreateOutputFiles(processedData, outputPath);

                Console.WriteLine();
                Console.WriteLine("========================================");
                Console.WriteLine("   Processing completed successfully!");
                Console.WriteLine("========================================");
                Console.WriteLine($"Output files created in: {outputPath}");

                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine("========================================");
                Console.WriteLine("             ERROR OCCURRED");
                Console.WriteLine("========================================");
                Console.WriteLine($"Error: {ex.Message}");
                return 1;
            }
        }

        static async Task CreateSampleInputFile(string filePath)
        {
            // Set EPPlus license for non-commercial use
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage();
            
            // Create Employees sheet
            var sheet1 = package.Workbook.Worksheets.Add("Employees");
            sheet1.Cells[1, 1].Value = "Employee ID";
            sheet1.Cells[1, 2].Value = "Employee Name";
            sheet1.Cells[1, 3].Value = "Department";
            sheet1.Cells[1, 4].Value = "Delegate Comments";
            sheet1.Cells[1, 5].Value = "Date";

            var sampleData = new[]
            {
                new { ID = "EMP001", Name = "John Doe", Dept = "IT", Comment = "New employee joining the team", Date = "2024-01-15" },
                new { ID = "EMP002", Name = "Jane Smith", Dept = "HR", Comment = "Update contact information", Date = "2024-01-16" },
                new { ID = "EMP003", Name = "Bob Johnson", Dept = "Finance", Comment = "Employee termination effective immediately", Date = "2024-01-17" },
                new { ID = "EMP004", Name = "Alice Brown", Dept = "Marketing", Comment = "Transfer to different location", Date = "2024-01-18" },
                new { ID = "EMP005", Name = "Charlie Wilson", Dept = "IT", Comment = "", Date = "2024-01-19" }
            };

            for (int i = 0; i < sampleData.Length; i++)
            {
                var row = i + 2;
                sheet1.Cells[row, 1].Value = sampleData[i].ID;
                sheet1.Cells[row, 2].Value = sampleData[i].Name;
                sheet1.Cells[row, 3].Value = sampleData[i].Dept;
                sheet1.Cells[row, 4].Value = sampleData[i].Comment;
                sheet1.Cells[row, 5].Value = sampleData[i].Date;
            }

            // Create Contractors sheet
            var sheet2 = package.Workbook.Worksheets.Add("Contractors");
            sheet2.Cells[1, 1].Value = "Employee ID";
            sheet2.Cells[1, 2].Value = "Employee Name";
            sheet2.Cells[1, 3].Value = "Department";
            sheet2.Cells[1, 4].Value = "Delegate Comments";
            sheet2.Cells[1, 5].Value = "Date";

            var contractorData = new[]
            {
                new { ID = "CON001", Name = "Mike Davis", Dept = "IT", Comment = "Add new contractor", Date = "2024-01-20" },
                new { ID = "CON002", Name = "Sarah Lee", Dept = "Design", Comment = "Modify contract terms", Date = "2024-01-21" },
                new { ID = "CON003", Name = "Tom Anderson", Dept = "Development", Comment = "Contract termination", Date = "2024-01-22" }
            };

            for (int i = 0; i < contractorData.Length; i++)
            {
                var row = i + 2;
                sheet2.Cells[row, 1].Value = contractorData[i].ID;
                sheet2.Cells[row, 2].Value = contractorData[i].Name;
                sheet2.Cells[row, 3].Value = contractorData[i].Dept;
                sheet2.Cells[row, 4].Value = contractorData[i].Comment;
                sheet2.Cells[row, 5].Value = contractorData[i].Date;
            }

            sheet1.Cells.AutoFitColumns();
            sheet2.Cells.AutoFitColumns();

            await package.SaveAsAsync(new FileInfo(filePath));
        }

        static async Task TrainModel(string modelPath)
        {
            // Default training data
            var trainingData = new List<CommentData>
            {
                // ADD samples
                new CommentData { Comment = "New employee joining", Action = "ADD" },
                new CommentData { Comment = "Add new user", Action = "ADD" },
                new CommentData { Comment = "Creating new account", Action = "ADD" },
                new CommentData { Comment = "Onboarding new staff", Action = "ADD" },
                new CommentData { Comment = "New hire", Action = "ADD" },
                new CommentData { Comment = "Fresh recruit", Action = "ADD" },
                new CommentData { Comment = "Addition to team", Action = "ADD" },
                new CommentData { Comment = "Add new contractor", Action = "ADD" },

                // UPDATE samples
                new CommentData { Comment = "Update employee details", Action = "UPDATE" },
                new CommentData { Comment = "Modify user information", Action = "UPDATE" },
                new CommentData { Comment = "Change department", Action = "UPDATE" },
                new CommentData { Comment = "Update contact info", Action = "UPDATE" },
                new CommentData { Comment = "Revise employee data", Action = "UPDATE" },
                new CommentData { Comment = "Edit user profile", Action = "UPDATE" },
                new CommentData { Comment = "Modify records", Action = "UPDATE" },
                new CommentData { Comment = "Modify contract terms", Action = "UPDATE" },

                // TERM samples
                new CommentData { Comment = "Employee termination", Action = "TERM" },
                new CommentData { Comment = "End of employment", Action = "TERM" },
                new CommentData { Comment = "Terminate user access", Action = "TERM" },
                new CommentData { Comment = "Remove from system", Action = "TERM" },
                new CommentData { Comment = "Resignation", Action = "TERM" },
                new CommentData { Comment = "Leaving company", Action = "TERM" },
                new CommentData { Comment = "Disable account", Action = "TERM" },
                new CommentData { Comment = "Contract termination", Action = "TERM" },
                new CommentData { Comment = "Employee termination effective immediately", Action = "TERM" },

                // OTHER samples
                new CommentData { Comment = "Transfer to different location", Action = "OTHER" },
                new CommentData { Comment = "Temporary suspension", Action = "OTHER" },
                new CommentData { Comment = "Leave of absence", Action = "OTHER" },
                new CommentData { Comment = "Role change pending", Action = "OTHER" },
                new CommentData { Comment = "Under review", Action = "OTHER" },
                new CommentData { Comment = "Special case", Action = "OTHER" },
                new CommentData { Comment = "Requires manual processing", Action = "OTHER" }
            };

            var data = _mlContext.Data.LoadFromEnumerable(trainingData);

            var pipeline = _mlContext.Transforms.Conversion.MapValueToKey(inputColumnName: "Action", outputColumnName: "Label")
                .Append(_mlContext.Transforms.Text.FeaturizeText(inputColumnName: "Comment", outputColumnName: "Features"))
                .Append(_mlContext.MulticlassClassification.Trainers.SdcaMaximumEntropy("Label", "Features"))
                .Append(_mlContext.Transforms.Conversion.MapKeyToValue("PredictedLabel"));

            _trainedModel = pipeline.Fit(data);

            // Save model
            using var fileStream = new FileStream(modelPath, FileMode.Create);
            _mlContext.Model.Save(_trainedModel, null, fileStream);

            Console.WriteLine("Model training completed and saved.");
        }

        static void LoadModel(string modelPath)
        {
            using var fileStream = new FileStream(modelPath, FileMode.Open, FileAccess.Read);
            _trainedModel = _mlContext.Model.Load(fileStream, out var modelInputSchema);
        }

        static async Task<List<ExcelRowData>> ReadExcelFile(string filePath)
        {
            // Set EPPlus license for non-commercial use
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            var allRowData = new List<ExcelRowData>();

            using var package = new ExcelPackage(new FileInfo(filePath));

            foreach (var worksheet in package.Workbook.Worksheets)
            {
                var startRow = worksheet.Dimension?.Start.Row ?? 1;
                var endRow = worksheet.Dimension?.End.Row ?? 1;
                var startCol = worksheet.Dimension?.Start.Column ?? 1;
                var endCol = worksheet.Dimension?.End.Column ?? 1;

                if (endRow <= startRow) continue;

                // Read header row
                var columnNames = new List<string>();
                for (int col = startCol; col <= endCol; col++)
                {
                    var headerValue = worksheet.Cells[startRow, col].Value?.ToString() ?? $"Column{col}";
                    columnNames.Add(headerValue);
                }

                // Find delegate comments column
                var delegateCommentsColumnIndex = columnNames.FindIndex(c => 
                    string.Equals(c, "Delegate Comments", StringComparison.OrdinalIgnoreCase));

                // Read data rows
                for (int row = startRow + 1; row <= endRow; row++)
                {
                    var rowItem = new ExcelRowData
                    {
                        SourceSheet = worksheet.Name,
                        RowNumber = row,
                        ColumnValues = new Dictionary<string, object>()
                    };

                    // Read all column values
                    for (int col = startCol; col <= endCol; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value ?? string.Empty;
                        var columnName = columnNames[col - startCol];
                        rowItem.ColumnValues[columnName] = cellValue;
                    }

                    // Set delegate comment
                    if (delegateCommentsColumnIndex >= 0)
                    {
                        var commentValue = worksheet.Cells[row, delegateCommentsColumnIndex + startCol].Value?.ToString();
                        rowItem.DelegateComment = commentValue ?? string.Empty;
                    }

                    allRowData.Add(rowItem);
                }
            }

            return allRowData;
        }

        static List<ExcelRowData> ProcessData(List<ExcelRowData> excelData)
        {
            if (_trainedModel == null) throw new InvalidOperationException("Model not loaded");

            var predictionEngine = _mlContext.Model.CreatePredictionEngine<CommentData, CommentPrediction>(_trainedModel);

            foreach (var row in excelData)
            {
                if (string.IsNullOrWhiteSpace(row.DelegateComment))
                {
                    row.PredictedAction = "ADD"; // Default as per requirements
                }
                else
                {
                    var input = new CommentData { Comment = row.DelegateComment };
                    var prediction = predictionEngine.Predict(input);
                    row.PredictedAction = prediction.PredictedAction;
                }
            }

            return excelData;
        }

        static void DisplaySummary(List<ExcelRowData> processedData)
        {
            Console.WriteLine();
            Console.WriteLine("=== Prediction Summary ===");
            
            var summary = processedData
                .GroupBy(d => d.PredictedAction)
                .OrderBy(g => g.Key)
                .ToDictionary(g => g.Key, g => g.Count());

            foreach (var kvp in summary)
            {
                Console.WriteLine($"{kvp.Key}: {kvp.Value} records");
            }

            Console.WriteLine($"Total: {processedData.Count} records");
            Console.WriteLine("==========================");
            Console.WriteLine();
        }

        static async Task CreateOutputFiles(List<ExcelRowData> processedData, string outputDirectory)
        {
            var actions = new[] { "ADD", "UPDATE", "TERM", "OTHER" };

            foreach (var action in actions)
            {
                var actionData = processedData.Where(d => 
                    string.Equals(d.PredictedAction, action, StringComparison.OrdinalIgnoreCase)).ToList();

                if (actionData.Any())
                {
                    var outputFileName = $"input_{action}.xlsx";
                    var outputPath = Path.Combine(outputDirectory, outputFileName);

                    await CreateSingleOutputFile(actionData, outputPath, action);
                    
                    Console.WriteLine($"Created: {outputFileName} with {actionData.Count} records");
                }
            }
        }

        static async Task CreateSingleOutputFile(List<ExcelRowData> data, string outputPath, string action)
        {
            // Set EPPlus license for non-commercial use
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add($"{action}_Records");

            if (!data.Any()) return;

            // Get all unique column names
            var allColumnNames = data
                .SelectMany(d => d.ColumnValues.Keys)
                .Distinct()
                .OrderBy(c => c)
                .ToList();

            // Add "Predicted Action" column
            if (!allColumnNames.Contains("Predicted Action"))
            {
                allColumnNames.Add("Predicted Action");
            }

            // Write header row
            for (int col = 0; col < allColumnNames.Count; col++)
            {
                worksheet.Cells[1, col + 1].Value = allColumnNames[col];
                worksheet.Cells[1, col + 1].Style.Font.Bold = true;
            }

            // Write data rows
            for (int row = 0; row < data.Count; row++)
            {
                var rowData = data[row];
                
                for (int col = 0; col < allColumnNames.Count; col++)
                {
                    var columnName = allColumnNames[col];
                    object cellValue;

                    if (columnName == "Predicted Action")
                    {
                        cellValue = rowData.PredictedAction;
                    }
                    else if (rowData.ColumnValues.TryGetValue(columnName, out var value))
                    {
                        cellValue = value;
                    }
                    else
                    {
                        cellValue = string.Empty;
                    }

                    worksheet.Cells[row + 2, col + 1].Value = cellValue;
                }
            }

            worksheet.Cells.AutoFitColumns();
            await package.SaveAsAsync(new FileInfo(outputPath));
        }
    }
}
