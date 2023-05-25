// See https://aka.ms/new-console-template for more information
using System.Globalization;
using System.Text;

const char COLUMN_SEPARATOR = ';';
const string INPUT_FILENAME = "Tuntiraportti.csv";
var INPUT_FILEPATH = AppDomain.CurrentDomain.BaseDirectory + "\\" + INPUT_FILENAME;
const string OUTPUT_FILENAME = "Tuntiraportti_projekteittain.csv";
var OUTPUT_FILEPATH = AppDomain.CurrentDomain.BaseDirectory + "\\" + OUTPUT_FILENAME;
const double WORKDAY_LENGTH = 7.25;

var values = File.ReadLines(INPUT_FILEPATH, Encoding.Latin1)
    .Skip(1)
    .Select(v => TuntikirjausData.FromCsv(v, COLUMN_SEPARATOR))
    .GroupBy(e => new {
        e.Toimintayksikkö,
        e.Project,
        e.Toiminto,
    })
    .Select(e => new OutputData(
        e.Key.Toimintayksikkö,
        e.First().ToimintayksikköName, // For some reason, the same toimintayksikkö can have multiple names
        e.Key.Toiminto,
        e.First().ToimintoName, // The same toiminto does not have multiple names, but just for consistency's sake it's not used as key
        e.Key.Project,
        e.First().ProjectName, // For some reason, the same project can have multiple names
        double.Round(e.ToList().Sum(e => e.AllocatedHours.TotalHours), 2),
        double.Round(e.ToList().Sum(e => e.AllocatedHours.TotalHours) / WORKDAY_LENGTH, 2))) // too lazy to not sum it twice
    .OrderBy(e => e.Project)
    .ToList();

var csvHeaderRow = "ToimintaYksikkö;ToimintaYksikköName;Toiminto;ToimintoName;Project;ProjectName;AllocatedHours;AllocatedDays" + "\n";
File.WriteAllText(OUTPUT_FILEPATH, csvHeaderRow + string.Join("\n", values), Encoding.Latin1);

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

    public static TuntikirjausData FromCsv(string csvLine, char columnSeparator) {
        var values = csvLine.Split(columnSeparator);
        var dailyValues = new TuntikirjausData(
            ContractNumber: values[0],
            Title: values[1],
            Date: DateTime.Parse(values[2], CultureInfo.GetCultureInfo("fi")),
            AllocatedHours: new TimeSpan(int.Parse(values[3].Split(":")[0]), int.Parse(values[3].Split(":")[1]), 0),
            Toimintayksikkö: values[4],
            ToimintayksikköName: values[5],
            Toiminto: values[6],
            ToimintoName: values[7],
            Project: values[8],
            ProjectName: values[9],
            Suorite: values[10],
            SuoriteName: values[11],
            Seuko1: values[12],
            Seuko1Name: values[13],
            Seuko2: values[14],
            Seuko2Name: values[15],
            Explanation: values[16],
            State: values[17]
        );
        return dailyValues;
    }
}