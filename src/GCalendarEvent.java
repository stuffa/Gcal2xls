import com.google.gdata.data.DateTime;

/**
 * @author Chris Martin
 *
 */


public class GCalendarEvent
{
  private final int MILISECONDS_IN_HOUR   = 60 * 60 * 1000;

  private String    calendarName;
  private String    task;
  private DateTime  starts;
  private DateTime  ends;
  private String    attendeeName;
  private String    attendeeEmail;
  private int       attendeeCount;
  private String    description;

  
  public GCalendarEvent()
  {
    calendarName     = "";
  	task             = "";
  	starts           = new DateTime();
  	ends             = new DateTime();
  	attendeeName     = "";
  	attendeeEmail    = "";
  	attendeeCount    = 0;
  	description      = "";
  }


  public String getCalendarName()
  {
      return calendarName;
  }
  public void setCalendarName(String calendarName)
  {
      this.calendarName = calendarName;
  }
 
  public String getTask()
  {
  	return task;
  }
  public void setTask(String task)
  {
    this.task = task;
  }

  public float getHours()
  {
      float hours = (float)(ends.getValue() - starts.getValue())/(MILISECONDS_IN_HOUR);
      return hours;
  }
  public DateTime getStarts()
  {
  	return this.starts;
  }
  public void setStarts(DateTime starts)
  {
    this.starts = starts;
  }

  public DateTime getEnds()
  {
  	return this.ends;
  }
  public void setEnds(DateTime ends)
  {
    this.ends = ends;
  }

  public String getAttendeeName()
  {
    return this.attendeeName;
  }
  public void setAttendeeName(String attendeeName)
  {
    this.attendeeName = attendeeName;
  }

  public String getAttendeeEmail()
  {
      return this.attendeeEmail;
  }
  public void setAttendeeEmail(String attendeeEmail)
  {
      this.attendeeEmail = attendeeEmail;
  }

  public int getAttendeeCount()
  {
      return this.attendeeCount;
  }
  public void setAttendeeCount(int attendeeCount)
  {
      this.attendeeCount = attendeeCount;
  }

  public String getDescription()
  {
      return this.description;
  }
  public void setDescription(String description)
  {
      this.description = description;
  }

  public int getTzShift()
  {
      return this.starts.getTzShift();
  }
  
}