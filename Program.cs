using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

using System.CommandLine;
using System.CommandLine.Parsing;

using OutLook = Microsoft.Office.Interop.Outlook;

class Program
{
    public static void Main(string[] args)
    {
        string thisappname = System.Diagnostics.Process.GetCurrentProcess().ProcessName;

        string main_description = "Outlook Bulk Calendar Json Importer\n\n" +
                                  "The program provides minimal input validation.\n" +
                                  "The data is provided as-is to the underlying Outlook API.";
                                
        var rootCommand = new RootCommand(main_description);
        var filenameArgument = new Argument<FileInfo>("jsonfile", "Json file containing list of meetings to import");
        var displayOption = new Option<bool>("--display", "Display the calendar windows for each meeting");
        var modalOption = new Option<bool>("--modal", "Modality for displaying windows (blocks between each meeting entry)");
        var testOption = new Option<bool>("--test", "Same as --delete. Hint: Test for errors or inspect meetings before saving or sending");
        var deleteOption = new Option<bool>("--delete", "Delete the meetings");
        var sendOption = new Option<bool>("--send", "Send the meeting request, requires typed input for confirmation");

        rootCommand.AddArgument(filenameArgument);
        rootCommand.AddOption(displayOption);
        rootCommand.AddOption(modalOption);
        rootCommand.AddOption(testOption);
        rootCommand.AddOption(deleteOption);
        rootCommand.AddOption(sendOption);

        rootCommand.AddValidator(
            commandResult =>
            {
                bool test = commandResult.Children.Any(sr => sr.Symbol is IdentifierSymbol id && id.HasAlias("--test"));
                bool delete = commandResult.Children.Any(sr => sr.Symbol is IdentifierSymbol id && id.HasAlias("--delete"));
                bool send = commandResult.Children.Any(sr => sr.Symbol is IdentifierSymbol id && id.HasAlias("--send"));

                if (send && (test || delete))
                    commandResult.ErrorMessage = "Option --send cannot be combined with --test or --delete";
            });

        rootCommand.SetHandler(
            (filename, display, modal, test, delete, send) =>
            {
                // If the send flag was set, request confirmation
                if (send)
                {
                    string? confirmation = null;

                    Console.Write("Type 'sEnD' and press enter to the meeting invites: ");
                    confirmation = Console.ReadLine();

                    if (confirmation != "sEnD")
                    {
                        Console.WriteLine("Your input '{0}' did not match.  Operation cancelled.", confirmation);
                        return;
                    }
                }

                // test & delete have the same meaning...the subroutine does not distinguish between test "or" delete [binary or]
                BulkyOlCal(filename, display, modal, test | delete, send);
            },
            filenameArgument, displayOption, modalOption, testOption, deleteOption, sendOption);

        rootCommand.InvokeAsync(args);
    }


    public static void BulkyOlCal(FileInfo fileinfo, bool display, bool modal, bool delete, bool send)
    {
        JsonSerializerOptions jsonOptions = new JsonSerializerOptions
        {
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            PropertyNameCaseInsensitive = true
        };
        Meeting[]? templates = null;

        try
        {
            string json = File.ReadAllText(fileinfo.FullName);
            templates = JsonSerializer.Deserialize<Meeting[]?>(json, jsonOptions);
        }
        catch (JsonException e)
        {
            // Line Numbers are 0-based.  Add one for human interpretation
            //string message = e.InnerException is FormatException ? e.InnerException.Message : e.Message;
            Console.Error.WriteLine("JSON Parse Error on Line {0}: {1}", (e.LineNumber + 1), e.InnerException?.Message ?? e.Message);
            return;
        }
        catch (Exception e)
        {
            Console.Error.WriteLine("Unexpected Error: ", e.Message);
            return;
        }

        if (templates == null) return;

        OutLook.Application olApp = new OutLook.Application();

        int i = 0;
        foreach (Meeting template in templates)
        {
            try
            {
                OutLook.AppointmentItem olAppt = (OutLook.AppointmentItem)olApp.CreateItem(OutLook.OlItemType.olAppointmentItem);

                if (template.Subject != null)
                    olAppt.Subject = template.Subject;
                if (template.Body != null)
                    olAppt.Body = template.Body;
                if (template.Location != null)
                    olAppt.Location = template.Location;
                if (template.Importance != null)
                    olAppt.Importance = template.GetOlImportance();
                if (template.Sensitivity != null)
                    olAppt.Sensitivity = template.GetOlSensitivity();
                if (template.BusyStatus != null)
                    olAppt.BusyStatus = template.GetOlBusyStatus();
                if (template.Start != null)
                    olAppt.Start = template.Start.Value;
                if (template.End != null)
                    olAppt.End = template.End.Value;
                if (template.Duration != null)
                    olAppt.Duration = template.Duration.Value;
                if (template.AllDayEvent != null)
                    olAppt.AllDayEvent = template.AllDayEvent.Value;
                if (template.ReminderSet != null)
                    olAppt.ReminderSet = template.ReminderSet.Value;
                if (template.ReminderMinutesBeforeStart != null)
                    olAppt.ReminderMinutesBeforeStart = template.ReminderMinutesBeforeStart.Value;
                if (template.ResponseRequested != null)
                    olAppt.ResponseRequested = template.ResponseRequested.Value;
    
                if (template.Recipients != null)
                {
                    // When recipients are included, an Appointment becomes a Meeting 
                    olAppt.MeetingStatus = OutLook.OlMeetingStatus.olMeeting;

                    foreach (Recipient recipient in template.Recipients)
                    {
                        OutLook.Recipient olRecipient = olAppt.Recipients.Add(recipient.Email);
                        olRecipient.Type = (int)recipient.GetOlMeetingRecipientType();
                    }
                }

                if (template.RecurrencePattern != null)
                {
                    OutLook.RecurrencePattern olRecur = olAppt.GetRecurrencePattern();

                    olRecur.RecurrenceType = template.RecurrencePattern.GetOlRecurrenceType();

                    if (template.RecurrencePattern.Interval != null)
                        olRecur.Interval = template.RecurrencePattern.Interval.Value;
                    if (template.RecurrencePattern.MonthOfYear != null)
                        olRecur.MonthOfYear = template.RecurrencePattern.MonthOfYear.Value;
                    if (template.RecurrencePattern.DayOfMonth != null)
                        olRecur.DayOfMonth = template.RecurrencePattern.DayOfMonth.Value;
                    if (template.RecurrencePattern.DayOfWeekMask != null)
                        olRecur.DayOfWeekMask = template.RecurrencePattern.GetOlDayOfWeekMask();
                    if (template.RecurrencePattern.Instance != null)
                        olRecur.Instance = template.RecurrencePattern.Instance.Value;

                    // Intnetionally set the pattern dates / occurrences / noenddate flag last
                    // After the pattern attributes are set, the Outlook api will re-calculate end date
                    // Number of occurrences based on the recurrence settings

                    if (template.RecurrencePattern.PatternStartDate != null)
                        olRecur.PatternStartDate = template.RecurrencePattern.PatternStartDate.Value;
                    if (template.RecurrencePattern.PatternEndDate != null)
                        olRecur.PatternEndDate = template.RecurrencePattern.PatternEndDate.Value;
                    if (template.RecurrencePattern.Occurences != null)
                        olRecur.Occurrences = template.RecurrencePattern.Occurences.Value;
                    if (template.RecurrencePattern.NoEndDate != null)
                        olRecur.NoEndDate = template.RecurrencePattern.NoEndDate.Value;

                    if (template.RecurrencePattern.StartTime != null)
                        olRecur.StartTime = template.RecurrencePattern.StartTime.Value;
                    if (template.RecurrencePattern.EndTime != null)
                        olRecur.EndTime = template.RecurrencePattern.EndTime.Value;
                    if (template.RecurrencePattern.Duration != null)
                        olRecur.Duration = template.RecurrencePattern.Duration.Value;
                }

                olAppt.Save();

                if (display)
                    olAppt.Display(modal);

                if (delete)
                    olAppt.Delete();
                else if (send)
                    olAppt.Send();
                else
                    olAppt.Close(OutLook.OlInspectorClose.olSave);
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("Error processing meeting #{0}: {1}", i+1, e.Message);
            }

            i++;
        }

    }
}

public class Meeting
{
    public Recipient[]? Recipients { get; set; }
    public string? Subject { get; set; }
    public string? Body { get; set; }
    public string? Location { get; set; }
    public string? Importance { get; set; }
    public string? Sensitivity { get; set; }
    public string? BusyStatus { get; set; }
    public DateTime? Start { get; set; }
    public DateTime? End { get; set; }
    public int? Duration { get; set; }
    public bool? AllDayEvent { get; set; }
    public bool? ReminderSet { get; set; }
    public int? ReminderMinutesBeforeStart { get; set; }
    public bool? ResponseRequested { get; set; }
    public RecurrencePattern? RecurrencePattern { get; set; }

    public OutLook.OlImportance GetOlImportance()
    {
        return Importance?.ToLower() switch
        {
            "low"    => OutLook.OlImportance.olImportanceLow,
            "normal" => OutLook.OlImportance.olImportanceNormal,
            "high"   => OutLook.OlImportance.olImportanceHigh,
            _        => throw new ArgumentException("Unexpected value for Importance")
        };
    }
    public OutLook.OlSensitivity GetOlSensitivity()
    {
        return Sensitivity?.ToLower() switch
        {
            "normal"       => OutLook.OlSensitivity.olNormal,
            "confidential" => OutLook.OlSensitivity.olConfidential,
            "personal"     => OutLook.OlSensitivity.olPersonal,
            "private"      => OutLook.OlSensitivity.olPrivate,
            _              => throw new ArgumentException("Unexpected value for Sensitivity")
        };
    }
    public OutLook.OlBusyStatus GetOlBusyStatus()
    {
        return BusyStatus?.ToLower() switch
        {
            "free"             => OutLook.OlBusyStatus.olFree,
            "tentative"        => OutLook.OlBusyStatus.olTentative,
            "busy"             => OutLook.OlBusyStatus.olBusy,
            "outofoffice"      => OutLook.OlBusyStatus.olOutOfOffice,
            "workingelsewhere" => OutLook.OlBusyStatus.olWorkingElsewhere,
            _                  => throw new ArgumentException("Unexpected value for BusyStatus")
        };
    }
}

public class Recipient
{
    public string? Email { get; set; }
    public string? Type { get; set; }
    public OutLook.OlMeetingRecipientType GetOlMeetingRecipientType()
    {
        return Type?.ToLower() switch
        {
            "organizer" => OutLook.OlMeetingRecipientType.olOrganizer,
            "required"  => OutLook.OlMeetingRecipientType.olRequired,
            "optional"  => OutLook.OlMeetingRecipientType.olOptional,
            "resource"  => OutLook.OlMeetingRecipientType.olResource,
            _           => throw new ArgumentException("Unexpected value for Recipient Type")
        };
    }
}
public class RecurrencePattern
{
    public string? RecurrenceType { get; set; }
    public DateTime? PatternStartDate { get; set; }
    public DateTime? PatternEndDate { get; set; }
    public int? Occurences { get; set; }
    public DateTime? StartTime { get; set; }
    public DateTime? EndTime { get; set; }
    public int? Duration { get; set; }
    public int? Interval { get; set; }
    public int? Instance { get; set; }
    public string[]? DayOfWeekMask { get; set; }
    public int? DayOfMonth { get; set; }
    public int? MonthOfYear { get; set; }
    public bool? NoEndDate { get; set; }

    public OutLook.OlRecurrenceType GetOlRecurrenceType()
    {
        return RecurrenceType?.ToLower() switch
        {
            "daily"      => OutLook.OlRecurrenceType.olRecursDaily,
            "weekly"     => OutLook.OlRecurrenceType.olRecursWeekly,
            "monthly"    => OutLook.OlRecurrenceType.olRecursMonthly,
            "monthlynth" => OutLook.OlRecurrenceType.olRecursMonthNth,
            "yearly"     => OutLook.OlRecurrenceType.olRecursYearly,
            "yearlynth"  => OutLook.OlRecurrenceType.olRecursYearNth,
            _            => throw new ArgumentException("Unexpected value for RecurrenceType")
        };
    }

    public OutLook.OlDaysOfWeek GetOlDayOfWeekMask()
    {
        OutLook.OlDaysOfWeek Mask = 0;

        if (DayOfWeekMask != null)
        {
            foreach (string Day in DayOfWeekMask)
            {
                Mask = Mask | Day?.ToLower() switch
                {
                    "sunday"    => OutLook.OlDaysOfWeek.olSunday,
                    "monday"    => OutLook.OlDaysOfWeek.olMonday,
                    "tuesday"   => OutLook.OlDaysOfWeek.olTuesday,
                    "wednesday" => OutLook.OlDaysOfWeek.olWednesday,
                    "thursday"  => OutLook.OlDaysOfWeek.olThursday,
                    "friday"    => OutLook.OlDaysOfWeek.olFriday,
                    "saturday"  => OutLook.OlDaysOfWeek.olSaturday,
                    "weekdays"  => OutLook.OlDaysOfWeek.olMonday | OutLook.OlDaysOfWeek.olTuesday | 
                                   OutLook.OlDaysOfWeek.olWednesday | OutLook.OlDaysOfWeek.olThursday |
                                   OutLook.OlDaysOfWeek.olFriday,
                    _           => throw new ArgumentException("Unexpected value for DayOfWeekMask")
                };
            }
        }

        return Mask;
    }
}