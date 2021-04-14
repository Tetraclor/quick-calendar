using CsvHelper;
using CsvHelper.Expressions;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace QuickCalendar
{
    class Program
    {
        static void Main(string[] args)
        {
            var cachedTimetibleFile = "timetable6sem_B18_191_1.csv";
            var sheetsID = "1TWBvHig9Smt44u_16uWp9aeWngT1YUURPHoXY8_GPw8";
            var column = "G";
            var calendarName = "6 Семестр";
            var endDate = new DateTime(2021, 5, 30);
            var loadFromRemoteSheet = false;
            var countPairs = 7;

            string[] timetable;

            if (loadFromRemoteSheet == false && File.Exists(cachedTimetibleFile))
            {
                timetable = File.ReadAllLines(cachedTimetibleFile);
            }
            else
            {
                timetable = LoadTimetableFromRemoteSheet(sheetsID, column);
                SaveTable(cachedTimetibleFile, timetable);
            }

            timetable = TimetableClear(timetable);
            timetable = TimetableSummaryShorten(timetable);
            
            var events = CreateEventsForRemoteCalendar(timetable, endDate, countPairs);

            CreateRemoteCalendar(events, calendarName);

            Console.WriteLine("End");
            Console.ReadKey();
        }

        static string[] LoadTimetableFromRemoteSheet(string sheetsID, string column)
        {
            column = column.ToUpper();
            var service = RemoteSheets.Auth();
            var result = RemoteSheets.Read(service, sheetsID, "Лист1", $"{column}:{column}");
            return result
                .Select(v => v.Any() ? v[0].ToString() : "").ToArray();
        }

        static void CreateRemoteCalendar(List<Event> events, string calendarName)
        {
            var calendarService = RemoteCalendar.Auth();
            var calendar = RemoteCalendar.CreateCalendar(calendarService, calendarName);

            foreach (var ev in events)
            {
                RemoteCalendar.AddEventToCalendar(calendarService, calendar.Id, ev);
            }
        }

        static List<Event> CreateEventsForRemoteCalendar(string[] timetable, DateTime endDateEvents, int CountPairs)
        {
            var aboveLine = new string[timetable.Length / 2];
            var belowLine = new string[timetable.Length / 2];

            for (var i = 0; i < timetable.Length / 2; i++)
            {
                aboveLine[i] = timetable[2 * i];
                belowLine[i] = timetable[2 * i + 1];
            }

            var timespans_even = CSVFormat.ReadFile<MySpan>("time_for_even_numbers");
            var above_events = TimetableToEventProcess(aboveLine, timespans_even, CountPairs, DateTime.Now.AddDays(7), endDateEvents);
            var below_events = TimetableToEventProcess(belowLine, timespans_even, CountPairs, DateTime.Now, endDateEvents);

            var events = new List<Event>();

            events.AddRange(above_events);
            events.AddRange(below_events);

            return events;
        }

        static List<Event> TimetableToEventProcess(string[] timetable, MySpan[] timespans, int countPairs, DateTime forDate, DateTime untilDate)
        {
            var events = new List<Event>();
            for(var i = 0; i < timetable.Length; i++)
            {
                if (timetable[i] == "" || string.IsNullOrEmpty(timetable[i])) continue;

                var day = (Day)(i / countPairs);
                var deltaDay = (int)MyDayEnumToDayOfWeekEnum(day) - (int)forDate.DayOfWeek + ((forDate - DateTime.Now).TotalDays);

                var ev = CreateEvent(timetable[i], timespans[i % countPairs]);
                ev.Start.DateTime = ev.Start.DateTime.Value.AddDays(deltaDay).AddHours(-1);
                ev.End.DateTime = ev.End.DateTime.Value.AddDays(deltaDay).AddHours(-1);
                ev = SetRecurrense(ev, day, untilDate);

                events.Add(ev);
            }
            return events;

            DayOfWeek MyDayEnumToDayOfWeekEnum(Day day)
            {
                return (DayOfWeek)(day == Day.SU ? 0 : ((int)day) + 1);
            }
        }

        static Event CreateEvent(string eventSummary, MySpan timeSpan)
        {
            return new Event()
            {
                Summary = eventSummary,
                Start = new EventDateTime()
                {
                    DateTime = DateTime.Parse(timeSpan.Start),
                    TimeZone = "Europe/Samara"
                },
                End = new EventDateTime()
                {
                    DateTime = DateTime.Parse(timeSpan.End),
                    TimeZone = "Europe/Samara"
                }
            };
        }

        enum Day
        {
            MO, // Monday
            TU, // Tuesday
            WE, // Wednesday
            TH, // Thursday
            FR, // Friday
            SA, // Saturday
            SU  // Sunday
        }

        static Event SetRecurrense(Event event_, Day byDay, DateTime until)
        {
            event_.Recurrence = new List<string>() 
            { 
                //$"EXDATE;VALUE=DATE:{start.ToString("yyyyMMdd")}",
                $"RRULE:FREQ=WEEKLY;INTERVAL=2;UNTIL={DateFormat(until)};BYDAY={byDay.ToString()}" 
            };
            return event_;

            string DateFormat(DateTime dataTime)
            {
                return dataTime.ToString("yyyyMMdd");
            }
        }

        static string[] TimetableClear(string[] timetable)
        {
            timetable = SkipHead(timetable);
            timetable = TrimRecords(timetable);

            return timetable;
            string[] TrimRecords(string[] timetable)
            {
                timetable = timetable.ToArray();
                for (var i = 0; i < timetable.Length; i++)
                {
                    timetable[i] = timetable[i].Trim('\"');
                    if (timetable[i].StartsWith("(")) //skip number in brackets
                        timetable[i] = timetable[i].Substring(4);
                }
                return timetable;
            }

            string[] SkipHead(string[] timetable)
            {
                return timetable.Skip(3).ToArray();
            }
        }

        static string[] TimetableSummaryShorten(string[] timetable) // timetable should be cleared
        {
            // (л/р) Компьютерная графика Левицкая Л.Н., ауд.3-204 => л/р 3-204 КГ Компьютерная графика Левицкая Л.Н.
            var result = new string[timetable.Length];
            for(var i = 0; i < timetable.Length; i++)
            {
                var record = timetable[i];

                if (record == "") continue;

                if (record.Contains("спорт"))
                {
                    result[i] = "Физра";
                    continue;
                }

                var splited = record.Split(' ').Where(v => v != "");
                var themeSplited = splited.Skip(2).TakeWhile(v => char.IsLower(v[0])).ToList();
                themeSplited.Insert(0, splited.Skip(1).First());

                var classType = splited.First();
                var auditory = splited.Last();
                var theme = string.Join(" ", themeSplited);
                var shortTheme = string.Concat(themeSplited.Select(v => char.ToUpper(v[0])));
                var teacher = string.Join(" ", splited.Skip(2).Where(v => char.IsUpper(v[0])));

                result[i] = $"{shortTheme} {classType} {auditory} {theme} {teacher}";
            }
            return result;
        }

        class MySpan
        {
            public string Start { get; set; }
            public string End { get; set; }
        }
        
        static void SaveTable(string filename, string[] timetable)
        {
            var table = timetable.Select(v => new List<string> () { v }).ToList();
            CSVFormat.WriteCSVTable(filename, table);
        }

        static void Tests()
        {
            //var credential = Auth();
            //var service = CreateCalendarService(credential);

            //RequestSample(service);
            //CreateSampleEvent(service);

            //var calendar = CreateSampleCalendar(service, "Test");

            // EventCSVFormat.CreateCSVFile("test_calendar", new List<Event>() { event_});

            //SheetsManager.Foo();
            //Console.WriteLine(new DateTime().ToString("t", new System.Globalization.CultureInfo("en-US")));

            var service = RemoteSheets.Auth();
            var table = RemoteSheets.Read(service, "1zcN4wBLvq_tnYrY7BfGjW2zBhcX_rET0DO6iBgEFZPw", "Лист2", "B2");

            CSVFormat.WriteCSVTable("table.csv", table);

            
        }

        static Event CreateSampleEvent(CalendarService service)
        {

            Event event_ = new Event()
            {
                Summary = "Google I/O 2015",
                Location = "800 Howard St., San Francisco, CA 94103",
                Description = "A chance to hear more about Google's developer products.",
                Start = new EventDateTime()
                {
                    DateTime = new DateTime(2020, 9, 1, 10, 0, 0),
                    TimeZone = "America/Los_Angeles"
                },
                End = new EventDateTime()
                {
                    DateTime = new DateTime(2020, 9, 1, 13, 0, 0),
                    TimeZone = "America/Los_Angeles"
                },

                Recurrence = new List<string>() { "RRULE:FREQ=DAILY;COUNT=2" },

                Attendees = new List<EventAttendee>() { new EventAttendee() { Email = "kdvikt@gmail.com" } },

                Reminders = new Event.RemindersData()
                {
                    UseDefault = false,
                    Overrides = new List<EventReminder>()
                    {
                        new EventReminder() { Method = "email", Minutes = 24 * 60 },
                        new EventReminder() { Method = "popup", Minutes = 10 }
                    }
                }
            };

            return event_;
        }

        static void RequestSample(CalendarService service)
    {
        // Define parameters of request.
        EventsResource.ListRequest request = service.Events.List("primary");
        request.TimeMin = DateTime.Now;
        request.ShowDeleted = false;
        request.SingleEvents = true;
        request.MaxResults = 10;
        request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

        // List events.
        Events events = request.Execute();
        Console.WriteLine("Upcoming events:");
        if (events.Items != null && events.Items.Count > 0)
        {
            foreach (var eventItem in events.Items)
            {
                string when = eventItem.Start.DateTime.ToString();
                if (String.IsNullOrEmpty(when))
                {
                    when = eventItem.Start.Date;
                }
                Console.WriteLine("{0} ({1})", eventItem.Summary, when);
            }
        }
        else
        {
            Console.WriteLine("No upcoming events found.");
        }
        Console.Read();
    }
    }
}
