import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.net.URL;
import java.net.URLDecoder;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import com.google.gdata.client.Query;
import com.google.gdata.client.calendar.CalendarQuery;
import com.google.gdata.client.calendar.CalendarService;
import com.google.gdata.data.DateTime;
import com.google.gdata.data.TextContent;
import com.google.gdata.data.calendar.CalendarEventEntry;
import com.google.gdata.data.calendar.CalendarEventFeed;
import com.google.gdata.data.calendar.EventWho;
import com.google.gdata.data.extensions.When;
import com.google.gdata.util.ServiceException;

/**
 * @author Chris Martin
 *
 */

public class GCalendarEvents extends LinkedList<GCalendarEvent>
{
    // NOTE: Google by default set a limit of 25 entries
    // If this was a feed rather than a Query we would use result.getNextLink().getHref()
    // to get the URL of the next page of results.
    // BUT.. queries are not feeds and they are not paged, so we are left with simply 
    // setting a large number for the MAX number of results/entries returned.
    // FIXME: Try and remove this hard limit if I can.
    private static final int	MAX_ENTRIES	= 10000;

    private static final long serialVersionUID	= 1L;
    private String googleApiId;
    
    private boolean excludeAllDayEvents		= false;
    private boolean seperateAttendees		= false;
    private boolean removeOwnerFromAttendees	= false;
    private boolean truncateDates		= false;
    
    private GContacts contacts			= null;
    

    // The base URL for a user's calendar metafeed (needs a username appended).
    // The URL for the event feed of the specified user's primary calendar.
    // (e.g. http://www.googe.com/feeds/calendar/calendar-id/private/full)
    private final String METAFEED_URL_BASE = "http://www.google.com/calendar/feeds/";
    private final String SINGLE_FEED_URL_SUFFIX = "/private/full";

    private final int MILISECONDS_IN_MINUTE = 60 * 1000;

    
    public GCalendarEvents(String googleApiId)
    {
	this.googleApiId = googleApiId;
    }

    
    // Options
    
    public void excludeAllDayEvents(boolean excludeAllDayEvents)
    {
        this.excludeAllDayEvents = excludeAllDayEvents;
    }
    
    public void seperateAttendees(boolean seperateAttenees)
    {
        this.seperateAttendees = seperateAttenees;
    }
    
    public void removeOwnerFromAttendees(boolean removeOwnerFromAttendees)
    {
        this.removeOwnerFromAttendees = removeOwnerFromAttendees;
    }

    public void truncateDates(boolean truncateDates)
    {
        this.truncateDates = truncateDates;
    }
    
    
    
    public void addEvents(String userName, char[] userPassword, GCalendars gCalendars, DateTime starts, DateTime ends) throws IOException, ServiceException
    {
		// Before we add any Events - first update the contacts so that we can use them to look up
		// appointments that only have email address
		contacts = new GContacts(googleApiId);
		contacts.update(userName, new String(userPassword));
		
		// This is the Google Calendar service we will be using
        CalendarService myService = new CalendarService(googleApiId);

        // authenticate with the user credentials
        myService.setUserCredentials(userName, new String(userPassword));

        for (GCalendar gCalendar : gCalendars)
        {
            if (gCalendar.isSelected())
                getEvents(myService, gCalendar, starts, ends);
        }
        
    }
    
    private void getEvents(CalendarService myService, GCalendar gCalendar, DateTime starts, DateTime ends) throws IOException, ServiceException
    {
        // Create the necessary URL objects.
        URL eventUrl = new URL(METAFEED_URL_BASE + gCalendar.getShortId() + SINGLE_FEED_URL_SUFFIX);

        // Set up the TZ offset to match that used in the calendar
        // this will ensure that we start correctly for each time-zone 
        starts.setTzShift(gCalendar.getTz().getOffset(starts.getValue())/MILISECONDS_IN_MINUTE);
        ends.setTzShift(gCalendar.getTz().getOffset(ends.getValue())/MILISECONDS_IN_MINUTE);

        // Build up the Query
        CalendarQuery myQuery = new CalendarQuery(eventUrl);
        myQuery.setMinimumStartTime(starts);
        myQuery.setMaximumStartTime(ends);
        myQuery.setMaxResults(MAX_ENTRIES);
        myQuery.addCustomParameter(new Query.CustomParameter("orderby",      "starttime"));
        myQuery.addCustomParameter(new Query.CustomParameter("sortorder",    "ascending"));
        myQuery.addCustomParameter(new Query.CustomParameter("singleevents", "true"));

        
        // Send the request and receive the response:
        CalendarEventFeed resultFeed = myService.query(myQuery, CalendarEventFeed.class);

        // process for each calendar entry received
        for ( CalendarEventEntry entry : resultFeed.getEntries())
        {
            // iterate over each time that the event occurred (i.e: recurring events)
            // Google will create a separate time entry for each instance.
            for (When when : entry.getTimes())
            {
                // Google also handles all-day events differently 
                // All day events can be detected by the following characteristics
                // (1) They have no time-shift
                // (2) There start and end dates are "DateOnly"

                if ( when.getStartTime().getTzShift() == null
                    && when.getStartTime().isDateOnly()
                    && when.getEndTime().isDateOnly() )
                {
                    // This is an all-day event
                    if ( excludeAllDayEvents)
                    {
                        // we are excluding them - skip it
                        continue;
                    }
                    else
                    {
                        // Displaying the All-Day events:  Do we truncate them
                        if (truncateDates)
                        {
                            // truncate the start time if necessary
                            
                            // NOTE: the Google event times are in GMT only
                            if ( when.getStartTime().getValue() < (starts.getValue() + (starts.getTzShift()*60*1000)) )
                                when.getStartTime().setValue(starts.getValue() + (starts.getTzShift()*60*1000));
                            
                            // truncate the end time if necessary
                            if ( when.getEndTime().getValue() > (ends.getValue() + (ends.getTzShift()*60*1000)) )
                                when.getEndTime().setValue(ends.getValue() + (ends.getTzShift()*60*1000));
                            
                            // check that we not left with an event with 0 duration
                            if ( when.getStartTime().getValue() == when.getEndTime().getValue() )
                                // skip it
                                continue;
                        }
                    }
                }
                else
                {
                    // Not an all day event:  Do we truncate it
                    if (truncateDates)
                    {
                        // truncate the start time if necessary
                        if ( when.getStartTime().getValue() < starts.getValue() )
                            when.getStartTime().setValue(starts.getValue());
                        
                        // truncate the end time if necessary
                        if ( when.getEndTime().getValue() > ends.getValue())
                            when.getEndTime().setValue(ends.getValue());
                        
                        // check that we have not created an event with 0 duration
                        if ( when.getStartTime().getValue() == when.getEndTime().getValue() )
                            // skip it
                            continue;
                    }
                }

                if (removeOwnerFromAttendees)
                {
                    // Google always adds the calendar (Name and ID) as an attendee
                    // iterate the list and remove the calendar attendee
                    // Note: I can't remove the entry while I iterate over the list
                    // so I'll make a list of entries to delete and delete after the loop
                    List<EventWho> attendeesToDelete = new ArrayList<EventWho>();
                    for ( EventWho attendee : entry.getParticipants() )
                        if (attendee.getEmail().compareToIgnoreCase(URLDecoder.decode(gCalendar.getShortId(), "UTF-8")) == 0)
                            attendeesToDelete.add(attendee);
                    
                    // Now iterate the delete list;
                    for ( EventWho attendee : attendeesToDelete )
                        entry.getParticipants().remove(attendee);
                }
                
                // Check the number of attendees to this event.
                // if there is none we still want to log it
                // Otherwise iterate over the attendee list
                if ( entry.getParticipants().size() == 0 )   
                {
                    // make the entry 
                    GCalendarEvent event = new GCalendarEvent();
                    event.setCalendarName(gCalendar.getName());
                    event.setTask(entry.getTitle().getPlainText());
                    event.setStarts(when.getStartTime());
                    event.setEnds(when.getEndTime());
                    event.setDescription(((TextContent)entry.getContent()).getContent().getPlainText());

                    // save each parsed entry
                    this.add(event);
                }
                else
                {
                    if (seperateAttendees)
                    {
                        for (EventWho attendee : entry.getParticipants())
                        {
                            // make the entry
                            GCalendarEvent event = new GCalendarEvent();
                            event.setCalendarName(gCalendar.getName());
                            event.setTask(entry.getTitle().getPlainText());
                            event.setStarts(when.getStartTime());
                            event.setEnds(when.getEndTime());
                            event.setAttendeeEmail(attendee.getEmail());
                            event.setAttendeeCount(entry.getParticipants().size());
                            event.setDescription(((TextContent)entry.getContent()).getContent().getPlainText());

                            // There is a problem with names not being populated in attendees
                            // so we will look up the name in the Contacts list.
                            // If the name is present, it MUST be an email address.
                            // In which case we will use the name from contacts

                            String attendeeName = attendee.getValueString();
                            String contactName = contacts.lookup(attendee.getEmail());
                            if ( contactName == null || contactName.isEmpty() )
                                event.setAttendeeName(attendeeName);
                            else
                                event.setAttendeeName(contactName);
                        	
                            // save the parsed entry
                            this.add(event);
                        }
                    }
                    else
                    {
                        String attendeeNames = null;
                        String attendeeEmails = null;
                        
                        // put all attendees together as a comma separated string
                        for (EventWho attendee : entry.getParticipants())
                        {
                            // There is a problem with names not being populated in attendees
                            // so we will look up the name in the Contacts list.
                            // If the name is present, it MUST be an email address.
                            // In which case we will use the name from contacts

                            String listName;
                            String attendeeName = attendee.getValueString();
                            String contactName = contacts.lookup(attendee.getEmail());
                            if ( contactName == null || contactName.isEmpty() )
                                listName = attendeeName;
                            else
                                listName = contactName;

                            // make a list of names
                            if (attendeeNames == null)
                                attendeeNames = listName ;
                            else
                                attendeeNames += ", " + listName;
                            
                            // make a list of email address
                            if (attendeeEmails == null)
                                attendeeEmails = attendee.getEmail();
                            else
                                attendeeEmails += ", " + attendee.getEmail();
                        }
                        
                        // make the entry
                        GCalendarEvent event = new GCalendarEvent();
                        event.setCalendarName(gCalendar.getName());
                        event.setTask(entry.getTitle().getPlainText());
                        event.setStarts(when.getStartTime());
                        event.setEnds(when.getEndTime());
                        event.setAttendeeName(attendeeNames);
                        event.setAttendeeEmail(attendeeEmails);
                        event.setAttendeeCount(entry.getParticipants().size());
                        event.setDescription(((TextContent)entry.getContent()).getContent().getPlainText());
                  
                        // save the entry
                        this.add(event);
                    }
                }
            }
        }
    }

    
    // Write all the Events to an Excel File
    public void writeXls(String fileName) throws RowsExceededException, WriteException, IOException
    {
        // add extension to the filename
        fileName += ".xls";
        
        // Create the spreadsheet
        WritableWorkbook workbook = Workbook.createWorkbook(new File(fileName));
        WritableSheet sheet = workbook.createSheet("The Sheet", 0);

        // set up the format
        WritableFont labelFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD, true);
        WritableCellFormat labelFormat = new WritableCellFormat(labelFont);

        // write the labels on the first row
        int col = 0;
        
        Label labelCalendar = new Label(col++, 0, Gcal2xls.res.getString("events.calendar.label"), labelFormat);
        sheet.addCell(labelCalendar);

        Label labelName = new Label(col++, 0, Gcal2xls.res.getString("events.attendee.label"), labelFormat);
        sheet.addCell(labelName);

        Label labelEmail = new Label(col++, 0, Gcal2xls.res.getString("events.email.label"), labelFormat);
        sheet.addCell(labelEmail);

        Label labelTask = new Label(col++, 0, Gcal2xls.res.getString("events.task.label"), labelFormat);
        sheet.addCell(labelTask);
        
        Label labelCount = new Label(col++, 0, Gcal2xls.res.getString("events.count.label"), labelFormat);
        sheet.addCell(labelCount);
        
        Label labelHours = new Label(col++, 0, Gcal2xls.res.getString("events.hours.label"), labelFormat);
        sheet.addCell(labelHours);

        Label labelStart = new Label(col++, 0, Gcal2xls.res.getString("events.start.label"), labelFormat);
        sheet.addCell(labelStart);

        Label labelEnds = new Label(col++, 0, Gcal2xls.res.getString("events.end.label"), labelFormat);
        sheet.addCell(labelEnds);

        Label labelDescription = new Label(col++, 0, Gcal2xls.res.getString("events.desc.label"), labelFormat);
        sheet.addCell(labelDescription);

        // leave a blank row
        int row = 2;

        for ( GCalendarEvent event : this)
        {
            col = 0;
            
            jxl.write.Label calendar = new jxl.write.Label(col++, row, event.getCalendarName());
            sheet.addCell(calendar);

            jxl.write.Label name = new jxl.write.Label(col++, row, event.getAttendeeName());
            sheet.addCell(name);

            jxl.write.Label email = new jxl.write.Label(col++, row, event.getAttendeeEmail());
            sheet.addCell(email);

            jxl.write.Label task = new jxl.write.Label(col++, row, event.getTask());
            sheet.addCell(task);

            jxl.write.Number count = new jxl.write.Number(col++, row, event.getAttendeeCount());
            sheet.addCell(count);

            jxl.write.Number hours = new jxl.write.Number(col++, row, event.getHours());
            sheet.addCell(hours);

            jxl.write.Label starts = new jxl.write.Label(col++, row, event.getStarts().toUiString());
            sheet.addCell(starts);

            // for all day events - display the end date nicely
            jxl.write.Label ends = null;
            if (event.getEnds().isDateOnly())
            {
                DateTime endDate = new DateTime(event.getEnds().getValue() -1);
                endDate.setDateOnly(true);
                ends = new jxl.write.Label(col++, row, endDate.toUiString());
            }
            else
            {
                ends = new jxl.write.Label(col++, row, event.getEnds().toUiString());
            }

            sheet.addCell(ends);

            jxl.write.Label description = new jxl.write.Label(col++, row, event.getDescription());
            sheet.addCell(description);

            row++;
        }

        workbook.write();
        workbook.close();

        return;
    }
    
    private static String encodeCVS(String str)
    {
      return "\"" + str.replace("\"", "\"\"") + "\"";  
    }

    
    public void writeCsv(String fileName) throws IOException
    {
        // add extension to the filename
        fileName += ".csv";
        
        FileWriter     csvFile;   // the actual file stream
        BufferedWriter csvWriter; // used to write
    
        csvFile = new FileWriter( new File(fileName) );
        csvWriter = new BufferedWriter(csvFile);

        // write the header
        csvWriter.write(
                encodeCVS(Gcal2xls.res.getString("events.calendar.label")) + "," +
                encodeCVS(Gcal2xls.res.getString("events.attendee.label")) + "," +
                encodeCVS(Gcal2xls.res.getString("events.email.label")) + "," +
                encodeCVS(Gcal2xls.res.getString("events.task.label")) + "," +
                encodeCVS(Gcal2xls.res.getString("events.count.label")) + "," +
                encodeCVS(Gcal2xls.res.getString("events.hours.label")) + "," +
                encodeCVS(Gcal2xls.res.getString("events.start.label")) + "," +
                encodeCVS(Gcal2xls.res.getString("events.end.label")) + "," +
                encodeCVS(Gcal2xls.res.getString("events.desc.label")) + "\n" );
        
        // now write events
        String row = null;
        for ( GCalendarEvent event : this)
        {
            row = encodeCVS(event.getCalendarName())  + "," + encodeCVS(event.getAttendeeName()) + "," +
            	  encodeCVS(event.getAttendeeEmail()) + "," + encodeCVS(event.getTask())         + "," +
            	  event.getAttendeeCount() +            "," + event.getHours()                   + "," +
            	  encodeCVS(event.getStarts().toUiString())  + ",";

            // for all day events - display the end date nicely
            if (event.getEnds().isDateOnly())
            {
                DateTime endDate = new DateTime(event.getEnds().getValue() -1);
                endDate.setDateOnly(true);
                row += encodeCVS(endDate.toUiString());
            }
            else
            {
                row += encodeCVS(event.getEnds().toUiString());
            }
            row += "," + encodeCVS(event.getDescription()) + "\n";
            
            csvWriter.write(row);
        }
        
        csvWriter.close();
        csvFile.close();

        return;
        
    }
}
