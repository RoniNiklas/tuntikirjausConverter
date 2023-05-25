// See https://aka.ms/new-console-template for more information
using ExcelDataReader;
using System.Globalization;
using System.Text;

const string INPUT_FILENAME = "Tuntiraportti.xlsx";
var INPUT_FILEPATH = AppDomain.CurrentDomain.BaseDirectory + "\\" + INPUT_FILENAME;
const string OUTPUT_FILENAME = "Tuntiraportti_projekteittain.csv";
var OUTPUT_FILEPATH = AppDomain.CurrentDomain.BaseDirectory + "\\" + OUTPUT_FILENAME;
const double WORKDAY_LENGTH = 7.25;

List<TuntikirjausData> tuntikirjausData = new();

using (var stream = File.Open(INPUT_FILEPATH, FileMode.Open, FileAccess.Read)) {
    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); // needed for .net core compatability
    using var reader = ExcelReaderFactory.CreateReader(stream);
    reader.Read(); // skip header row
    while (reader.Read()) {
        tuntikirjausData.Add(TuntikirjausData.FromReader(reader));
    }
}

var groupedValues = tuntikirjausData
    .GroupBy(e => new {
        e.Toimintayksikkö,
        e.Project,
        e.Toiminto,
    })
    .Select(e => new OutputData(
        e.Key.Toimintayksikkö,
        e.First().ToimintayksikköName, // For some reason, the same toimintayksikkö can have multiple names (basically just typoes of what it's supposed to be)
        e.Key.Toiminto,
        e.First().ToimintoName, // The same toiminto does not have multiple names, but just for consistency's sake it's not used as key
        e.Key.Project,
        e.First().ProjectName, // For some reason, the same project can have multiple names (basically just typoes of what it's supposed to be)
        double.Round(e.ToList().Sum(e => e.AllocatedHours.TotalHours), 2),
        double.Round(e.ToList().Sum(e => e.AllocatedHours.TotalHours) / WORKDAY_LENGTH, 2))) // too lazy to not sum it twice
    .OrderBy(e => e.Project)
    .ToList();

var csvHeaderRow = "ToimintaYksikkö;ToimintaYksikköName;Toiminto;ToimintoName;Project;ProjectName;AllocatedHours;AllocatedDays" + "\n";
File.WriteAllText(OUTPUT_FILEPATH, csvHeaderRow + string.Join("\n", groupedValues), Encoding.Latin1);

// End of program 

public record OutputData(string ToimintaYksikkö, string ToimintaYksikköName, string Toiminto, string ToimintoName, string Project, string ProjectName, double AllocatedHours, double AllocatedDays) {
    public override string ToString() {
        return $"{ToimintaYksikkö};{ToimintaYksikköName};{Toiminto};{ToimintoName};{Project};{ProjectName};{AllocatedHours};{AllocatedDays}";
    }
}

public record TuntikirjausData(
    string ContractNumber,
    string Title,
    DateTime Date,
    TimeSpan AllocatedHours,
    string Toimintayksikkö,
    string ToimintayksikköName,
    string Toiminto,
    string ToimintoName,
    string Project,
    string ProjectName,
    string Suorite,
    string SuoriteName,
    string Seuko1, // always empty
    string Seuko1Name,  // always empty
    string Seuko2,  // always empty 
    string Seuko2Name, // always empty
    string Explanation, // always empty
    string State) {
    public static TuntikirjausData FromReader(in IExcelDataReader reader) {
        var dailyValues = new TuntikirjausData(
            ContractNumber: reader.FromCell(0),
            Title: reader.FromCell(1),
            Date: DateTime.Parse(reader.FromCell(2), CultureInfo.GetCultureInfo("fi")),
            AllocatedHours: new TimeSpan(int.Parse(reader.FromCell(3).Split(":")[0]), int.Parse(reader.FromCell(3).Split(":")[1]), 0),
            Toimintayksikkö: reader.FromCell(4),
            ToimintayksikköName: reader.FromCell(5),
            Toiminto: reader.FromCell(6),
            ToimintoName: reader.FromCell(7),
            Project: reader.FromCell(8),
            ProjectName: reader.FromCell(9),
            Suorite: reader.FromCell(10),
            SuoriteName: reader.FromCell(11),
            Seuko1: reader.FromCell(12),
            Seuko1Name: reader.FromCell(13),
            Seuko2: reader.FromCell(14),
            Seuko2Name: reader.FromCell(15),
            Explanation: reader.FromCell(16),
            State: reader.FromCell(17)
        );
        return dailyValues;
    }
}

public static class IExcelDataReaderExtensions {
    public static string FromCell(this IExcelDataReader reader, in int column) {
        try {
            return reader.GetString(column);
        }
        catch (Exception e) {
            return ""; // an empty cell throws an exception which sucks so we'll just catch it and give back an empty string.
        }
    }
}