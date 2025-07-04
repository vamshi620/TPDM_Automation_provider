# TPDM Automation Provider

A .NET Core console application that uses ML.NET to process Excel files and predict actions based on delegate comments.

## Features

- **ML.NET Model**: Trains a machine learning model to predict actions (ADD, UPDATE, TERM, OTHER) based on delegate comments
- **Excel Processing**: Reads multiple sheets from Excel files using EPPlus
- **Automatic Prediction**: Processes "Delegate Comments" column and predicts appropriate actions
- **Output Generation**: Creates 4 separate Excel files for each action type
- **Default Handling**: Treats rows with no delegate comments as ADD operations

## Requirements

- .NET 8.0 or later
- Input Excel files with "Delegate Comments" column
- Non-commercial use (EPPlus license)

## Installation

1. Clone the repository
2. Navigate to the TPDMAutomationProvider directory
3. Run `dotnet restore` to install dependencies
4. Run `dotnet build` to build the application

## Usage

```bash
dotnet run
```

The application will:
1. Create a sample input file if none exists
2. Train or load the ML model
3. Process the input Excel file
4. Generate 4 output Excel files (one for each action type)

## Configuration

The application uses the following default paths:
- Input file: `./TestData/input.xlsx`
- Output directory: `./Output/`
- ML model: `./MLModels/comment_classifier.zip`

## Sample Data

The application creates sample data with:
- Employee records with various delegate comments
- Contractor records with different comment types
- Examples of ADD, UPDATE, TERM, and OTHER scenarios

## ML Model Training

The model is trained on sample data including:
- **ADD**: "New employee joining", "Add new user", "Creating new account"
- **UPDATE**: "Update employee details", "Modify user information", "Change department"
- **TERM**: "Employee termination", "End of employment", "Remove from system"
- **OTHER**: "Transfer to different location", "Temporary suspension", "Leave of absence"

## Output Files

The application generates 4 Excel files:
- `input_ADD.xlsx` - Records requiring addition
- `input_UPDATE.xlsx` - Records requiring updates
- `input_TERM.xlsx` - Records requiring termination
- `input_OTHER.xlsx` - Records requiring other actions

Each output file contains:
- All original columns from the input data
- A "Predicted Action" column showing the ML prediction
- Only records matching the specific action type

## Architecture

### Data Models
- `CommentData`: Input/training data for ML model
- `CommentPrediction`: ML model prediction results
- `ExcelRowData`: Excel row data with predictions

### Key Components
- **ML Training**: Uses SdcaMaximumEntropy for multiclass classification
- **Excel Processing**: Reads/writes Excel files with EPPlus
- **Prediction Engine**: Processes comments and predicts actions

## Example Output

```
========================================
   TPDM Automation Provider v1.0
========================================

Creating sample input file...
Training ML model...
Processing Excel file...
Making predictions...

=== Prediction Summary ===
ADD: 3 records
UPDATE: 2 records
TERM: 2 records
OTHER: 1 records
Total: 8 records
==========================

Creating output files...
Created: input_ADD.xlsx with 3 records
Created: input_UPDATE.xlsx with 2 records
Created: input_TERM.xlsx with 2 records
Created: input_OTHER.xlsx with 1 records

Processing completed successfully!
========================================
```

## Dependencies

- **Microsoft.ML**: Machine learning framework
- **EPPlus**: Excel file processing (version 6.2.10 for licensing compatibility)
- **Microsoft.Extensions.Configuration**: Configuration management
- **Microsoft.Extensions.Configuration.Json**: JSON configuration support
- **Microsoft.Extensions.Configuration.Binder**: Configuration binding

## License

This application uses EPPlus under non-commercial license terms. For commercial use, please ensure proper EPPlus licensing.

## Error Handling

The application includes comprehensive error handling for:
- Missing input files (creates sample data)
- Excel file processing errors
- ML model training/loading issues
- File system operations

## Performance

- Processes Excel files with multiple sheets efficiently
- Batch prediction capabilities for large datasets
- Persistent ML model storage to avoid retraining