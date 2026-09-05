using System.IO;

// Define the base path for the backups
// ImpactIQ v3: the orchestrator passes the base folder explicitly (IMPACTIQ_BASE); the working directory stays the fallback.
string baseFolderPath = Environment.GetEnvironmentVariable("IMPACTIQ_BASE");
if (string.IsNullOrEmpty(baseFolderPath)) baseFolderPath = Directory.GetCurrentDirectory();
var addedPath = System.IO.Path.Combine(baseFolderPath, "Model Backups");

// Dynamically find the latest-dated folder
string[] folders = System.IO.Directory.GetDirectories(addedPath);
string latestFolder = null;
DateTime latestDate = DateTime.MinValue;

foreach (string folder in folders)
{
    string folderName = System.IO.Path.GetFileName(folder);
    DateTime folderDate;

    if (DateTime.TryParseExact(folderName, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out folderDate))
    {
        if (folderDate > latestDate)
        {
            latestDate = folderDate;
            latestFolder = folder;
        }
    }
}

// ImpactIQ v3: IMPACTIQ_DATE_FOLDER names the run folder to write to (bypasses the latest-folder heuristic); IMPACTIQ_REPORT_DATE the ModelAsOfDate string.
string iqDateFolder = Environment.GetEnvironmentVariable("IMPACTIQ_DATE_FOLDER");
if (!string.IsNullOrEmpty(iqDateFolder) && System.IO.Directory.Exists(iqDateFolder))
{
    latestFolder = iqDateFolder;
    DateTime iqFolderDate;
    if (DateTime.TryParseExact(System.IO.Path.GetFileName(iqDateFolder), "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out iqFolderDate)) latestDate = iqFolderDate;
}
string iqReportDate = Environment.GetEnvironmentVariable("IMPACTIQ_REPORT_DATE");

// Use the latest-dated folder, or fallback to today's date if no valid folder is found
var currentDateStr = !string.IsNullOrEmpty(iqReportDate) ? iqReportDate : (latestFolder != null && latestDate != DateTime.MinValue ? latestDate.ToString("yyyy-MM-dd") : DateTime.Now.ToString("yyyy-MM-dd"));

// Create the folder path for the backup
var dateFolderPath = latestFolder ?? System.IO.Path.Combine(addedPath, currentDateStr);
if (!System.IO.Directory.Exists(dateFolderPath))
{
    System.IO.Directory.CreateDirectory(dateFolderPath);
}

// Retrieve the model name
var modelName = Model.Database.Name;
var modelID = Model.Database.ID;

// Initialize the StringBuilder for the CSV content
var sb = new System.Text.StringBuilder();
sb.AppendLine("ObjectName,ObjectType,DependsOn,DependsOnType,ModelAsOfDate,ModelName,ModelID");

// ImpactIQ v3 (audit X1-03): quote every field and double embedded quotes, like the Model Detail script does.
Func<dynamic, string> FormatField = (field) =>
{
    if (field == null) { return "\"\""; }
    string text = field.ToString();
    if (string.IsNullOrEmpty(text)) { return "\"\""; }
    return "\"" + text.Replace("\"", "\"\"") + "\"";
};

// ===============================
//   MEASURES
// ===============================
foreach (var table in Model.Tables)
{
    foreach (var measure in table.Measures)
    {
        var dependencies = measure.DependsOn;

        foreach (var dependency in dependencies)
        {
            sb.AppendLine(string.Join(",", new string[] {
                                         FormatField(measure.Name),
                                         FormatField("Measure"),
                                         FormatField(dependency.Key.DaxObjectFullName),
                                         FormatField(dependency.Key.ObjectType.ToString()),
                                         FormatField(currentDateStr),
                                         FormatField(modelName),
                                         FormatField(modelID) }));
        }
    }
}

// ===============================
//   CALCULATED COLUMNS
// ===============================
foreach (var table in Model.Tables)
{
    foreach (var calcCol in table.CalculatedColumns)
    {
        var dependencies = calcCol.DependsOn;

        foreach (var dependency in dependencies)
        {
            sb.AppendLine(string.Join(",", new string[] {
                                         FormatField(calcCol.Name),
                                         FormatField("CalculatedColumn"),
                                         FormatField(dependency.Key.DaxObjectFullName),
                                         FormatField(dependency.Key.ObjectType.ToString()),
                                         FormatField(currentDateStr),
                                         FormatField(modelName),
                                         FormatField(modelID) }));
        }
    }
}

// ===============================
//   CALCULATION ITEMS
// ===============================
foreach (var calcGroup in Model.CalculationGroups)
{
    foreach (var calcItem in calcGroup.CalculationItems)
    {
        var dependencies = calcItem.DependsOn;

        foreach (var dependency in dependencies)
        {
            sb.AppendLine(string.Join(",", new string[] {
                FormatField(calcItem.Name),
                FormatField("CalculationItem"),
                FormatField(dependency.Key.DaxObjectFullName),
                FormatField(dependency.Key.ObjectType.ToString()),
                FormatField(currentDateStr),
                FormatField(modelName),
                FormatField(modelID) }));
        }
    }
}

// Write the file
var filePath = System.IO.Path.Combine(dateFolderPath, modelName + "_MD.csv");
System.IO.File.WriteAllText(filePath, sb.ToString());
