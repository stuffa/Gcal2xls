import java.io.IOException;

import java.net.URL;

import java.util.Collections;
import java.util.TimeZone;
import java.util.Vector;

import com.google.gdata.client.calendar.CalendarService;
import com.google.gdata.data.calendar.CalendarEntry;
import com.google.gdata.data.calendar.CalendarFeed;
import com.google.gdata.util.ServiceException;

/**
 * @author Chris Martin
 *
 */
public class GCalendars extends Vector<GCalendar>
{
    private static final long serialVersionUID = 3L;

    // This is the URL that is used to query for the lost of Calendars from Google
    private final String CALENDAR_LIST_URL = "http://www.google.com/calendar/feeds/default/allcalendars/full";

    private String	googleApiId;
    private boolean	includeHiddenCalendars = false;

    
    public GCalendars(String googleApiId)
    {
	this.googleApiId = googleApiId;
    }

    public GCalendars (String userName, String userPassword, String googleApiId) throws IOException, ServiceException
    {
	this.googleApiId = googleApiId;
        update(userName, userPassword);
    }
    
    public void setIncludeHiddenCalendars(boolean includeHiddenCalendars)
    {
	this.includeHiddenCalendars = includeHiddenCalendars;
    }

    /**
     * @param Google User Name
     * @param Google User Password 
     * @throws ServiceException 
     * @throws IOException 
     */
    public void update(String userName, String userPassword) throws IOException, ServiceException
    {
        // first - mark all calendar entries as eligible for deletion
        for (GCalendar cal : this)
            cal.setToBeDeleated();
        
        // This is the Google calendar service object.
        CalendarService calService = new CalendarService(googleApiId);
        
        // Attempt to get the list of calendars for this user from Google

        // Create the  URL objects.
        URL calendarListUrl = new URL(CALENDAR_LIST_URL);

        // authenticate with the user credentials
        calService.setUserCredentials(userName, userPassword);

        while (calendarListUrl != null)
        {
            // make the query to get the list    
            CalendarFeed calendarFeed = calService.getFeed(calendarListUrl, CalendarFeed.class);
            
            // iterate over the feed and build a Calendar list
            for (CalendarEntry entry : calendarFeed.getEntries())
            {
        	if (includeHiddenCalendars == false)
        	    if ( entry.getHidden().getValue().equalsIgnoreCase("true") )
        		continue;
        	
                // Create a new Calendar for this entry
                GCalendar newCal = new  GCalendar(entry.getId(), entry.getTitle().getPlainText(), TimeZone.getTimeZone(entry.getTimeZone().getValue()));
                
                // we want to preserve the selected state of existing calendar entries
                // first see if we have a matching calendar
                for (GCalendar oldCal : this)
                {
                    if (oldCal.getId().compareTo(newCal.getId()) == 0)
                    {
                        // make the selected state of the new entry the same as the old
                        newCal.setSelected(oldCal.isSelected());
                        // remove the current calendar item
                        remove(oldCal);
                        break;
                    }
                }
                
                // add the new Calendar entry
                this.add(newCal);
            }
            
            // check that there are no more "extra" calendars to fetch
            if (calendarFeed.getNextLink() == null)
        	calendarListUrl = null;
            else
        	calendarListUrl = new URL(calendarFeed.getNextLink().getHref());
        }
        
        // now remove any calendar entries that marked for deletion
        // First we make a new list of the entries to be deleted.
        //  I do this to avoid problems if all calendars are to be deleted
        GCalendars deletedList = new GCalendars(googleApiId);
        for (GCalendar cal : this)
            if (cal.isToBeDeleted())
                deletedList.add(cal);
        
        // now delete the deleatedList
        for (GCalendar cal : deletedList)
            remove(cal);
        
        // Now sort them
        Collections.sort(this);

        return;
     }
}
