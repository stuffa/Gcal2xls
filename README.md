# Gcal2xls


This utility will extract events from your Google calendar (or Google Apps Calendar) and save the information in a spreadsheet - currently either Microsoft Excel (xls) or as Comma Separated Values (CVS).

To run the application you will require Java 1.5 or higher.  Depending on your system, you may be either able to launch it by clicking on the jar file icon, or with the following command line:

    java -jar Gcal2xls

#### Hidden Calendars
If you have a large number of calenders, some may be hidden.  This option will allow them all to be visable in the list of calenders.

#### Calenders
You can export the events fom multiple calendars.  Select the ones you want to export events from

#### Date Range
Enter the start and end dates for the events that you wantto export

#### File
Select the file format and location that the file should be written to

#### Options
##### Remove calendar owner from event attendees
This option will remove the calendar owner from the list of attendees.  

##### List events by attendee
This will group the events by attendees.  If there where multiple attendes the event will be display ed from each of them

##### Exclude all-day events
This option will exclude all day events.  These are usually birthdays, anniversiarys etc.

##### Truncate dates to query range
The default action will list all evenst that fall within the given date range, and report the full lenghth of the event.
This option will truncate the reported event to the range specified.  This is usyally usefull for billing time over a specified period.
