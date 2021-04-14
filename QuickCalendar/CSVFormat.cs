using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.Configuration.Attributes;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3.Data;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace QuickCalendar
{
    public class CSVFormat
    {
        public static void WriteCSVEvents(string fileName, IEnumerable<Event> events)
        {
            using (var writer = new StreamWriter(fileName))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(events.Select(v => new EventObjectMap(v)));
            }
        }

        public static T[] ReadFile<T>(string filename)
        {
            using (var reader = new StreamReader(filename))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                return csv.GetRecords<T>().ToArray();
            }
        }

        public static IEnumerable<IEnumerable<string>> ReadTable(string fileName)
        {
            return new List<List<string>>();
        }

        public static void WriteCSVTable(string fileName, IEnumerable<IEnumerable<object>> data)
        {
            using (var writer = new StreamWriter(fileName))
            using (var csv = new CsvSerializer(writer, CultureInfo.InvariantCulture))
            {
                foreach (var line in data)
                {
                    var result = line.Select(v => Format(v.ToString())).ToArray();
                    csv.Write(result);
                    csv.WriteLine();
                }
            }

            string Format(string value)
            {
                value = value.Replace('\n', ' ');
                return value.Contains(",") ? $"\"{value}\"" : value;
            }
        }

        class EventObjectMap
        {
            [Name("Subject")]
            public string Subject { get; set; }

            [Name("Start date")]
            public string StartDate { get; set; }

            [Name("Start time")]
            public string StartTime { get; set; }

            [Name("End date")]
            public string EndDate { get; set; }

            [Name("End time")]
            public string EndTime { get; set; }

            public EventObjectMap(Event event_)
            {
                var startDate = event_.Start.DateTime.Value;
                var endDate = event_.End.DateTime.Value;


                Subject = event_.Summary;
                StartDate = $"{startDate.Month}/{startDate.Day}/{startDate.Year}";
                StartTime = startDate.ToString("t", new System.Globalization.CultureInfo("en-US"));
                EndDate = $"{endDate.Month}/{endDate.Day}/{endDate.Year}";
                EndTime = endDate.ToString("t", new System.Globalization.CultureInfo("en-US"));
            }
        }
    }

    //public class EventMap : ClassMap<Event>
    //{
    //    public EventMap()
    //    {
    //        Map(v => v.Summary).Index(0).Name("Subject");
    //        Map(v => v.Start.Date).Index(1).Name("Start date");
    //        Map(v => $"{v.Start.DateTime.Value.Hour}:{v.Start.DateTime.Value.Minute}").Index(2).Name("Start time");
    //        Map(v => $"{v.End.DateTime.Value.Hour}:{v.End.DateTime.Value.Minute}").Index(3).Name("End time");
    //    }
    //}
}

