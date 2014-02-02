import javax.swing.AbstractButton;
import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JRadioButton;
import javax.swing.JScrollPane;
import javax.swing.JSeparator;
import javax.swing.JTextField;
import javax.swing.ListCellRenderer;
import javax.swing.ListSelectionModel;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.border.Border;
import javax.swing.border.EmptyBorder;

import net.miginfocom.swing.MigLayout;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.awt.Color;
import java.awt.Component;
import java.awt.Container;
import java.awt.Cursor;
import java.awt.Font;
import java.awt.GridLayout;
import java.awt.HeadlessException;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;

import com.google.gdata.data.DateTime;
import com.google.gdata.util.AuthenticationException;
import com.google.gdata.util.ServiceException;
import com.toedter.calendar.JDateChooser;

/**
 * Extract Events from Google Calendar and save to s spreadsheet 
 *  
 * @author Chris Martin
 */

public class Gcal2xls extends JFrame implements ActionListener, WindowListener, MouseListener, KeyListener
{
    private static final long serialVersionUID = 3L;
    private static final String version = "2.02";
    private static final String googleApiId = "cc.martin.gcal2xls-" + version;
        
    private JTextField username;
    private JPasswordField password;
      
    private JList<GCalendar>       calendarChooser;
    private JScrollPane gCalenderscrollPane;
     
    private JDateChooser startsDateChooser;
    private JDateChooser endsDateChooser;

    private JTextField   dirName;
    private JButton	 dirChooserBtn;
    private JFileChooser dirChooser; 
    
    private ButtonGroup  conversionFormat;
    private JRadioButton doXls;
    private JRadioButton doCsv;
    private JPanel 	 formatChooser;
    
//    private JCheckBox checkAll;

    private JCheckBox includeHiddenCalendars;
    private JCheckBox seperateAttendees;
    private JCheckBox excludeAllDayEvents;
    private JCheckBox removeOwnerFromAttendees;
    private JCheckBox truncateDates;
    
    private JButton clearCalList;
    private JButton updateCalList;
    private JButton extractDates;
      
    private final String DATE_FORMAT_PATTERN = "yyyy-MM-dd";
    private SimpleDateFormat dateFormat = new SimpleDateFormat(DATE_FORMAT_PATTERN);;
    
    private GCalendars gCalendars = new GCalendars(googleApiId);
  
   
      
    /**
     * The constructor.
     */
    public Gcal2xls()
    {
        Container container = getContentPane();
        container.setLayout(new MigLayout("", "[][grow][::135]","[][][][][][grow][]"));
        
        // Try to set the look and feel
        try
	{
	    UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
	}
        catch (Exception e)
        {
            // do nothing
        }
        SwingUtilities.updateComponentTreeUI(container);
        

        //initialise the main panel
        setTitle("Gcal2xls v" + version  );
        setSize(500,700);
        setResizable(true);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        addWindowListener(this);

        // Login detail separator
        container.add(new JHeading("Authentication"), "split, span, gaptop 10");
        container.add(new JSeparator(),             "growx, wrap, gaptop 10");

        // username
        username = new JTextField();
        username.setToolTipText("e.g. user@domain.com");
        container.add(new JLabel("User Name:"));
        container.add(username, "growx");
        container.add(new JHint("<html>e.g. user@domain.com</html>"), "wrap");

        // password
      	password = new JPasswordField();
        password.setToolTipText("Your Google Gmail or Google Apps password");
        container.add(new JLabel("Password:"));
        container.add(password, "growx");
        container.add(new JHint("<html>Your Google (Gmail or Apps) password</html>"), "wrap");

        // calendar detail separator
        container.add(new JHeading("Calendars"), "split, span, gaptop 10");
        container.add(new JSeparator(),          "growx, wrap, gaptop 10");
        
        //The clearCalList  Button
//        clearCalList = new JButton("Clear List");
//        clearCalList.addActionListener(this);
//        container.add(clearCalList, "skip 1, split 2");

        // CheckAll button
//        checkAll = new JCheckBox("Check All");
//        container.add(checkAll, "skip 1, split 2");

        // Select Hidden Calendars
        includeHiddenCalendars = new JCheckBox("Include hidden calendars");
        includeHiddenCalendars.setToolTipText("Include calendars that you have disabled/hidden in your calendar view");
        container.add(includeHiddenCalendars, "skip 1, gapright push, wrap");

        // The updateCalList Button
        updateCalList = new JButton("Update List");
        updateCalList.addActionListener(this);
        container.add(updateCalList, "skip 1, gapright push, wrap");

        // The list of gCalenders
        // Note we haven't got a list of gCalenders yet so they can't be displayed
        // This is an empty list that we add to latter
        calendarChooser = new JList<GCalendar>(gCalendars);
        calendarChooser.setToolTipText("This list is extracted from from your Google calendar settings");
        CheckListCellRenderer renderer = new CheckListCellRenderer();
        calendarChooser.setCellRenderer(renderer);
        calendarChooser.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        calendarChooser.addMouseListener(this);
        calendarChooser.addKeyListener(this);
        gCalenderscrollPane = new JScrollPane();
        gCalenderscrollPane.getViewport().add(calendarChooser);
        container.add(new JLabel("Calenders:"));
        container.add(gCalenderscrollPane, "growx, growy");
        container.add(new JHint("<html>This list is extracted from from your Google calendar settings</html>"), "wrap");
        
        //Start Date Choosers
      	startsDateChooser = new JDateChooser(new Date(), DATE_FORMAT_PATTERN);
      	container.add(new JLabel("Starts:"));
        container.add(startsDateChooser, "growx");
      	container.add(new JHint("<html>yyyy-mm-dd</html>"), "wrap");

      	//End Date Choosers
      	endsDateChooser = new JDateChooser(new Date(), DATE_FORMAT_PATTERN);
      	container.add(new JLabel("Ends:"));
        container.add(endsDateChooser,"growx");
      	container.add(new JHint("<html>yyyy-mm-dd inclusive</html>"), "wrap");

        // calendar detail separator
        container.add(new JHeading("Extracted File"), "split, span, gaptop 10");
        container.add(new JSeparator(),               "growx, wrap, gaptop 10");
      	
      	//Setup the radio buttons
        doXls = new JRadioButton("XLS (Microsoft Excel)");
        doXls.setActionCommand("xls");
        doXls.setToolTipText("Save the data in XLS (Microsoft Excel) Format");
        doXls.setSelected(true);

        doCsv = new JRadioButton("CSV (Comma Seperated Values)");
        doCsv.setActionCommand("csv");
        doCsv.setToolTipText("Save the data in CSV (Comma Seperated Values) Format");
        
        //Group the radio buttons.
        conversionFormat = new ButtonGroup();
        conversionFormat.add(doXls);
        conversionFormat.add(doCsv);
        
        //Put the radio buttons in a column
        formatChooser = new JPanel(new GridLayout(0, 1));
        formatChooser.add(doXls);
        formatChooser.add(doCsv);

        // add the radio buttons to the layout
        container.add(new JLabel("Format:"));
        container.add(formatChooser, "wrap");


      	// Output File Location
        container.add(new JLabel("Location:"));
      	
      	dirName = new JTextField(); 
        container.add(dirName, "growx, split 2");

      	// Folder Button
        dirChooserBtn = new JButton();
        dirChooserBtn.setActionCommand("create");
        dirChooserBtn.setIcon(UIManager.getDefaults( ).getIcon( "FileView.directoryIcon" ));
     	dirChooserBtn.addActionListener(this);
        // The Folder Chooser
        dirChooser = new JFileChooser();
        dirChooser.setDialogTitle("Export File Location");
        dirChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        dirChooser.setAcceptAllFileFilterUsed(false);
        // set the default 
        String homeDir = System.getProperty("user.home");
        String curDir = System.getProperty("user.dir");
        if (homeDir != null & homeDir.length() > 0)
        {
            dirChooser.setCurrentDirectory(new java.io.File(homeDir));
            dirName.setText(homeDir);
        }
        else
        {
            if (curDir != null & curDir.length() > 0)
            {
                dirChooser.setCurrentDirectory(new java.io.File(curDir));
                dirName.setText(curDir);
            }
        }
        container.add(dirChooserBtn);
      	container.add(new JHint("<html>Location where files are written</html>"), "wrap");

        
        // The options
        removeOwnerFromAttendees = new JCheckBox("Remove calendar owner from event attendees");
        removeOwnerFromAttendees.setToolTipText("Dont list the calender (owner) as an attendee, by default the calender is always listed");
        container.add(removeOwnerFromAttendees, "skip 1, spanx 2, wrap");

        seperateAttendees = new JCheckBox("List events by attendee");
        seperateAttendees.setToolTipText("Events with multiple attendies will be repeated/split. Once for each attendee");
        container.add(seperateAttendees, "skip 1, spanx 2, wrap");

        excludeAllDayEvents = new JCheckBox("Exclude all-day events");
        excludeAllDayEvents.setToolTipText("Exclude all day events eg: birthdays, from the ouput");
        container.add(excludeAllDayEvents, "skip 1, spanx 2, wrap");
        
        truncateDates = new JCheckBox("Truncate dates to query range");
        truncateDates.setToolTipText("Events that span outside the selected range will be truncated to the selected dates");
        container.add(truncateDates, "skip 1, spanx 2, wrap");
        
        
        //The Extract Button
        extractDates = new JButton("Extract");
        extractDates.addActionListener(this);
        container.add(extractDates, "gaptop 10, skip 1, align right");

        
        // populate with saved data
        this.loadUserInfo();

        setVisible(true);
    }

    /*
     * The ActionListener
     */  
    public void actionPerformed(ActionEvent event)
    {
      Object object = event.getSource();
      
      if ( object == extractDates )
      {
        extract_actionPerformed(event);
        return;
      }
      
      if ( object == updateCalList )
      {
          populateCalendarList();
          return;
      }

      if ( object == clearCalList )
      {
	  gCalendars.clear();
          calendarChooser.setListData(gCalendars);
          return;
      }
      
      if ( object == dirChooserBtn )
      {
          dirChooser.setCurrentDirectory(new java.io.File(dirName.getText()));
          if (dirChooser.showOpenDialog(getContentPane()) == JFileChooser.APPROVE_OPTION)
          { 
              dirName.setText(dirChooser.getSelectedFile().getPath());
          }
      }
      return;

    }
  
    
    /*
     * WindowClosingListener 
     */
    public void windowActivated(WindowEvent e)
    {
        return;
    }
    public void windowClosed(WindowEvent e)
    {
        return;
    }
    public void windowClosing(WindowEvent e)
    {
        saveUserInfo();
        return;
    }
    public void windowDeactivated(WindowEvent e)
    {
        return;
    }
    public void windowDeiconified(WindowEvent e)
    {
        return;
    }
    public void windowIconified(WindowEvent e)
    {
        return;
    }
    public void windowOpened(WindowEvent e)
    {
        return;
    }
    
    /*
     * Mouse Listener
     */
    
    public void mouseClicked(MouseEvent e)
    {
            doCheck();
    }

    public void mousePressed(MouseEvent e) {}

    public void mouseReleased(MouseEvent e) {}

    public void mouseEntered(MouseEvent e) {}

    public void mouseExited(MouseEvent e) {}
    
    /*
     * KeyListener
     */

    public void keyPressed(KeyEvent e)
    {
        if (e.getKeyChar() == ' ')
            doCheck();
    }

    public void keyTyped(KeyEvent e) {}
    
    public void keyReleased(KeyEvent e) {}


    /*
     * Check 
     */
    
    protected void doCheck()
    {
        int index = calendarChooser.getSelectedIndex();
        if (index < 0)
            return;
        GCalendar gCalendar = (GCalendar)calendarChooser.getModel().getElementAt(index);
        gCalendar.invertSelection();
        calendarChooser.repaint();
    }

      
    private void populateCalendarList()
    {
        // Check password and name fields to see if we have any.
        if ( (username.getText().length() != 0) && (password.getPassword().length != 0) )
        {
            // make a busy cursor
            setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
            
           // fetch the list of gCalenders
           try
           {
               gCalendars.setIncludeHiddenCalendars(includeHiddenCalendars.isSelected());
               gCalendars.update(username.getText(), new String(password.getPassword()));
               calendarChooser.setListData(gCalendars);
           }
           catch (AuthenticationException e)
           {
               displayError("Authentication Error:", "Invalid crededtials\nPlease check User Name & Password");
               return;
           }
           catch (IOException e)
           {
               displayError("IO Exception:", e.getLocalizedMessage());
               return;
           }
           catch (ServiceException e)
           {
               displayError("Service Exception:", e.getLocalizedMessage());
               return;
           }
           finally
           {
               setCursor(Cursor.getDefaultCursor());
           }
        }
        else
        {
            displayError("Authentication Error:", "Missing credentials\nPlease check User Name & Password");
            return;
        }
    }
    
    public static void main(String args[])
    {
        new Gcal2xls();
    }
  
    private void extract_actionPerformed(ActionEvent event)
    {
      	String userName 	= username.getText();
        String userPassword = new String (password.getPassword());

        DateTime starts;
        DateTime ends;
        
        // The dates fetched from the GUI may contain the current time.
        // This has to be removed.
        
        // No timezone has been specified so this is UTC/GMT date&time
        // correct timezone compensation is applied once we have the default TZ for the calendar
        // it may be different for each calendar.
        
        try
        {
            // Get the date making sure that we don't have any time associated with it.
            Calendar sDate = startsDateChooser.getCalendar();
            sDate.set(Calendar.HOUR, 0);
            sDate.set(Calendar.MINUTE, 0);
            sDate.set(Calendar.SECOND, 0);
            sDate.set(Calendar.MILLISECOND, 0);
            sDate.set(Calendar.AM_PM, Calendar.AM);
            
            starts = new DateTime (sDate.getTimeInMillis());
            starts.setDateOnly(true);
        }
        catch (NullPointerException e)
        {
            displayError("Start Date Error:", "Invalid start date");
            return;
        }
        
        try
        {
            Calendar eDate = endsDateChooser.getCalendar();
            // remove any time
            eDate.set(Calendar.HOUR, 0);
            eDate.set(Calendar.MINUTE, 0);
            eDate.set(Calendar.SECOND, 0);
            eDate.set(Calendar.MILLISECOND, 0);
            eDate.set(Calendar.AM_PM, Calendar.AM);
            // now move forward a day
            eDate.add(Calendar.DAY_OF_MONTH, 1);

            ends = new DateTime (eDate.getTimeInMillis());
            ends.setDateOnly(true);
        }
        catch (NullPointerException e)
        {
            displayError("End Date Error:", "Invalid end date");
            return;
        }
    
    	// check that the dates are correct
    	if (ends.getValue() < starts.getValue())
    	{
            displayError("Date Order Error:", "Please check dates.\nThe start date must be equal or ealier that the end date");
            return;
    	}
    	
    	
        // fetch the events data from the calendar
    	GEvents events = new GEvents(googleApiId);
    	events.seperateAttendees(seperateAttendees.isSelected());
    	events.excludeAllDayEvents(excludeAllDayEvents.isSelected());
    	events.removeOwnerFromAttendees(removeOwnerFromAttendees.isSelected());
        events.truncateDates(truncateDates.isSelected());

        // make a busy cursor
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        try
        {
            events.addEvents(userName, userPassword, gCalendars, starts, ends);
        }
        catch (AuthenticationException e)
        {
            displayError("Authentication Error:", "Please check User Name & Password");
            return;
        }
        catch (IOException e)
        {
            displayError("IO Error", e.getLocalizedMessage());
            return;
        }
        catch (ServiceException e)
        {
            displayError("Service Error", e.getLocalizedMessage());
            return;
        }
        finally
        {
            setCursor(Cursor.getDefaultCursor());
        }
    
    	if (events == null)
    	{
    	    displayError("NULL Erorr", "Events retruned was NULL");
    	    return;
    	}
    	
       // The name of the spreadsheet
       String fileName = dirName.getText() + System.getProperty( "file.separator") + userName + " TimeSheet From " + dateFormat.format(startsDateChooser.getDate()) + " To " + dateFormat.format(endsDateChooser.getDate());
    
       // write the data in the selected format
        try
        {
            if (doXls.isSelected())
                events.writeXls(fileName);
            else if (doCsv.isSelected())
                events.writeCsv(fileName);
            else
                displayError("Error", "Unsupported Format");
        }
        catch (IOException e)
        {
            // ToDo: display a error message
            displayError("IO Error", e.getLocalizedMessage());
            return;
        }
        catch (jxl.write.WriteException e)
        {
            // ToDo: display a error message
            displayError("IO Error", e.getLocalizedMessage());
            return;
        }
    
       	displayMessage("Successfully Converted", "File was created successfully\n" + events.size() + " records");
    }

    private void displayMessage(String title, String message)
    {
        try
        {
            JOptionPane.showMessageDialog(this, message, title, JOptionPane.INFORMATION_MESSAGE);
        }
        catch (HeadlessException e)
        {
            e.printStackTrace();
        }
    }
  
    private void displayError(String title, String message)
    {
        try
        {
            JOptionPane.showMessageDialog(this, message, title, JOptionPane.ERROR_MESSAGE);
        }
        catch (HeadlessException e)
        {
            e.printStackTrace();
        }
    }

    private void saveUserInfo()
    {
       try
        {
            FileOutputStream settingsFile = new FileOutputStream(this.getClass().getName() + ".dat");
            ObjectOutputStream settingsStream = new ObjectOutputStream(settingsFile);

            settingsStream.writeObject(username.getText());
            settingsStream.writeObject(dirName.getText());
            settingsStream.writeObject(gCalendars);
            settingsStream.writeObject(Collections.list( conversionFormat.getElements()));
            settingsStream.writeObject(removeOwnerFromAttendees.isSelected());
            settingsStream.writeObject(seperateAttendees.isSelected());
            settingsStream.writeObject(excludeAllDayEvents.isSelected());
            settingsStream.writeObject(truncateDates.isSelected());
            settingsStream.writeObject(includeHiddenCalendars.isSelected());

            settingsStream.close();
            settingsFile.close();
        }
        catch (Exception e)
        {
            displayError("IO Error", "Unable to save parameter file\n" + e.getLocalizedMessage());
        }

        return;
    }
  
    @SuppressWarnings("unchecked")
    private void loadUserInfo()
    {

        try
        {
            Boolean selected = false;
            
            FileInputStream settingsFile = new FileInputStream(this.getClass().getName() + ".dat");
            ObjectInputStream settingsStream = new ObjectInputStream(settingsFile);
    
            // restore the user name
            username.setText((String) settingsStream.readObject());
            // restore the location
            dirName.setText((String) settingsStream.readObject());
            dirChooser.setCurrentDirectory(new java.io.File(dirName.getText()));
            // restore the calendar
            gCalendars = (GCalendars) settingsStream.readObject();
            if (gCalendars == null)
                gCalendars = new GCalendars(googleApiId);
            calendarChooser.setListData(gCalendars);
            // get the format buttons
            List<AbstractButton> buttonList = (List<AbstractButton>) settingsStream.readObject();
            List<AbstractButton> radioList = Collections.list(conversionFormat.getElements());
            Iterator<AbstractButton> radioIterator = radioList.iterator();
            for (AbstractButton button : buttonList)
            {   
                AbstractButton radio = radioIterator.next();
                if( button.isSelected() )
                    radio.setSelected(true);
            }

            selected = (Boolean)settingsStream.readObject();
            removeOwnerFromAttendees.setSelected(selected);
            selected = (Boolean)settingsStream.readObject();
            seperateAttendees.setSelected(selected);
            selected = (Boolean)settingsStream.readObject();
            excludeAllDayEvents.setSelected(selected);
            selected = (Boolean)settingsStream.readObject();
            truncateDates.setSelected(selected);
            selected = (Boolean)settingsStream.readObject();
            includeHiddenCalendars.setSelected(selected);
            
            settingsStream.close();
            settingsFile.close();
        }
        catch (Exception e)
        {
            // recover
            if (gCalendars == null)
            {
                gCalendars = new GCalendars(googleApiId);
                calendarChooser.setListData(gCalendars);
            }
        }
        finally
        {
            if (username.getText().length() == 0)
                // set the focus on the user name
                username.requestFocus();
            else
                // put the focus on the password
                password.requestFocus();
        }

        return;
    }

}

class CheckListCellRenderer extends JCheckBox implements ListCellRenderer
{
    private static final long serialVersionUID = 1L;
    protected static Border m_noFocusBorder = new EmptyBorder(1, 1, 1, 1);

    public CheckListCellRenderer()
    {
        super();
        setOpaque(true);
        setBorder(m_noFocusBorder);
    }

    public Component getListCellRendererComponent(JList list,
        Object value, int index, boolean isSelected, boolean cellHasFocus)
    {
	// Set the display value
        setText(value.toString());
        // Cast the value passed so that we can call the isSelected method to determine
        // if the value was already selected, and set the JList appropriately
        setSelected(((GCalendar)value).isSelected());

        setBackground(isSelected ? list.getSelectionBackground() : list.getBackground());
        setForeground(isSelected ? list.getSelectionForeground() : list.getForeground());
        return this;
    }
}

class JHint extends JLabel
{
    private static final long serialVersionUID = 1L;
    
    JHint(String hint)
    {
	super(hint);
	this.setForeground(Color.gray);
	Font defaultFont = this.getFont();
	this.setFont(new Font(defaultFont.getFamily(), defaultFont.getStyle(), (int)(defaultFont.getSize2D()*0.9)));
    }
}

class JHeading extends JLabel
{
    private static final long serialVersionUID = 1L;
    
    JHeading(String hint)
    {
	super(hint);
	Font defaultFont = this.getFont();
	this.setFont(new Font(defaultFont.getFamily(), Font.BOLD, defaultFont.getSize()));
    }
}
