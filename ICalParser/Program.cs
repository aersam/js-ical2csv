using Ical.Net;
using Ical.Net.CalendarComponents;
using System;
using System.Linq;

namespace ICalParser
{
    class Program
    {
        static void Main(string[] args)
        {
            var wc = new System.Net.WebClient();
            wc.DownloadFile("https://calendar.google.com/calendar/ical/vosgbbgk9130p9ollrkhm5vnq4%40group.calendar.google.com/public/basic.ics", "Termine.ics");
            string ical = System.IO.File.ReadAllText(@"Termine.ics");
            var cal = Calendar.Load(ical);
            var dtStart = new DateTime(DateTime.Now.Year, 1, 1);//Always current year
            using (var output = new System.IO.StreamWriter("output.csv", false, System.Text.Encoding.UTF8))
            using (var csvHelper = new CsvHelper.CsvWriter(output))
            {
                
                csvHelper.WriteField("Datum");
                csvHelper.WriteField("Was");
                csvHelper.WriteField("Beschreibung");
                csvHelper.WriteField("Input");
                csvHelper.NextRecord();

                foreach (var child in cal.Children.Cast<CalendarEvent>()
                    .Where(c => c.Start.Value > dtStart)
                    .OrderBy(c => c.Start.Value))
                {
                    string[] desc = child.Description
                        .Replace("\r\n", "\n")    
                        .Split('\n');
                    string input = desc.FirstOrDefault(d => d.StartsWith("Input:"));
                    string realDesc = string.Join("\r\n", desc.Where(d => d != input));
                    csvHelper.WriteField(child.Start.Value.ToString("dd.MM.yyyy"));
                    csvHelper.WriteField(child.Summary);
                    csvHelper.WriteField(realDesc);
                    csvHelper.WriteField(input?.Substring("Input:".Length)?.Trim());
                    csvHelper.NextRecord();
                }
            }
            Console.WriteLine("Done. Press enter");
            Console.ReadKey();
        }
    }
}
