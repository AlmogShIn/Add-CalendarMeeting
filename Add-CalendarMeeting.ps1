function Add-CalendarMeeting {
    [CmdletBinding()]    
    param(
        # The meeting subject
        [Parameter(
            Mandatory = $True,
            HelpMessage = "Please provide a subject of you calendar evnite.")]
        [Alias('sub')] 
        [string] $Subject,

        # The meeting body - description of the meeting
        [Parameter(
            Mandatory = $True,
            HelpMessage = "Please provide a description of you calendar evnite."
        )]
        [Alias('bod')]
        [string] $Body,

        # The meeting Required participants
        [Parameter(
            Mandatory = $True,
            HelpMessage = "Please provide a participants for you calendar evnite."
        )]
        [Alias('ReqAtt')]
        [string[]] $ReqAttendees,

        # The meeting Optinal participants
        [Parameter(
            Mandatory = $True,
            HelpMessage = "Please provide a participants for you calendar evnite."
        )]
        [Alias('OptAtt')]
        [string[]] $OptAttendees,

        # Meeting location
        [String] $Location = "Teams",

        # Importance Parameter . 0 - low Importance 1 - High importance
        [Parameter(
            HelpMessage = "0 - low Importance 1 - High importance"
        )]
        [int] $Importance = 0,

        #Set Reminder parameter
        [bool] $EnableReminder = $True,

        # Meeting Remainder 
        [int] $ReminderMinBefore,

        # Meeting Duration
        [int] $Duration = 25,

        # Meeting Start Time
        [datetime] $MeetingStart = (Get-Date)
    )
    
    Begin 
    {
        # Create Outlook Object
        $outlookApplication = New-Object -ComObject 'Outlook.Application'

        # Create Appointment Item
        $newCalendarItem = $outlookApplication.CreateItem('olAppointmentItem')

        if(-not $Subject){
        
        }
    }
    Process 
    { 
        $newCalendarItem.meetingstatus = [Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeeting
        $newCalendarItem.Subject = $Subject
        $newCalendarItem.Body = $Body
        $newCalendarItem.Location = $Location
        $newCalendarItem.ReminderSet = $EnableReminder
        $newCalendarItem.ReminderMinutesBeforeStart = $ReminderMinBefore
        $newCalendarItem.Importance = $Importance
        $newCalendarItem.Start = $MeetingStart
        $newCalendarItem.Duration = $Duration

        foreach ($attend in $ReqAttendees) {
            $newCalendarItem.Recipients.Add($attend)
        }

        foreach ($attend in $OptAttendees) {
            $newCalendarItem.OptionalAttendees.Add($attend)
        }

        $newCalendarItem.Send()
        $newCalendarItem.Save()
    }
}

$newMeetingTime = Get-Date -Date 'Monday, August 08, 2022 2:27:29 PM'
$Recipients = @('almog.shtaigmann@intel.com', 'tamir.zitman@intel.com')

Add-CalendarMeeting  -Body "I wish its work!" -ReqAttendee $Recipients -MeetingStart $newMeetingTime





    


