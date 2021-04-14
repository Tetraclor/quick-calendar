using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace QuickCalendar
{
    class RemoteCalendar
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/calendar-dotnet-quickstart.json
        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "Google Calendar API .NET Quickstart";

        public static CalendarService Auth()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }


            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return service;
        }

        public static void AddEventToCalendar(CalendarService service, string calendarId, Event event_)
        {
            event_ = service.Events.Insert(event_, calendarId).Execute();
            Console.WriteLine("Event created: " + event_.HtmlLink);
        }

        public static Calendar CreateCalendar(CalendarService service, string summary)
        {
            var calendar = new Calendar()
            {
                Summary = summary
            };

            calendar = service.Calendars.Insert(calendar).Execute();

            Console.WriteLine($"Calendar created: {calendar.Summary}");

            return calendar;
        }
    }
}
