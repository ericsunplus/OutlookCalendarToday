import win32com.client
import datetime as dt
from prettytable import PrettyTable

outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('Mapi')
calendar = outlook.getDefaultFolder(9).Items        # calendar
calendar.IncludeRecurrences = True
calendar.Sort("[Start]")
begin = dt.datetime.today()
end = begin + dt.timedelta(days = 1)
restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
targetEmployee = 'Eric Sun'
calendar = calendar.Restrict(restriction)

t = PrettyTable(['Time', 'Subject', 'Organizer', 'Required', 'Accepted'])
t.hrules = True
t.align['Subject'] = 'l'

for item in calendar:
    time = item.StartInStartTimeZone.strftime('%H:%M');
    subject = item.Subject
    orgnizer = item.Organizer
    required = 'Required' if targetEmployee in item.RequiredAttendees else ''
    accepted = 'Accepted' if item.ResponseStatus == 3 else ''
    t.add_row([time, subject, orgnizer, required, accepted])
print(t)

#
# olResponseNone            0   The appointment is a simple appointment and does not require a response.
# olResponseOrganized       1   The AppointmentItem is on the Organizer's calendar or the recipient is the Organizer of the meeting.
# olResponseTentative       2   Meeting tentatively accepted.
# olResponseAccepted        3   Meeting accepted.
# olResponseDeclined        4   Meeting declined.
# olResponseNotResponded    5   Recipient has not responded.
#
