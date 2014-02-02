import java.io.Serializable;
import java.util.TimeZone;

/**
 * @author Chris Martin
 *
 */


public class GCalendar implements Comparable<GCalendar>, Serializable 
{
    
    /**
     * 
     */
    private static final long serialVersionUID = 1L;
    private String      name;
    private String      id;
    private TimeZone    tz;
    private boolean     selected;
    private boolean     toBeDeleted;
    
    public GCalendar (String id, String name, TimeZone tz)
    {
        this.id     = id;
        this.name   = name;
        this.tz     = tz;
        selected    = true;
        toBeDeleted = false;
    }
    
    public boolean isToBeDeleted()
    {
        return toBeDeleted;
    }
    
    public void setToBeDeleated()
    {
        toBeDeleted = true;
    }
    
    public boolean isSelected()
    {
        return selected;
    }
    
    public void setSelected(boolean selected)
    {
        this.selected = selected;
    }
    
    public void invertSelection()
    {
        selected = !selected;
    }
    public String getName()
    {
        return name;
    }
    
    public TimeZone getTz()
    {
        return tz;
    }
    
    public String getId()
    {
        return id;
    }
    
    public String getShortId()
    {
        // get the basename of the ID url
        return id.substring(id.lastIndexOf('/')+1);
    }
    
    public String toString()
    {
        return name;
    }

    // we only compare to our own type
    public int compareTo(GCalendar gCalendar)
    {
        return name.compareToIgnoreCase(gCalendar.name);
    }

}
