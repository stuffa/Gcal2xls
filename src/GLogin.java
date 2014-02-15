import java.awt.Container;
import java.awt.Cursor;
import java.awt.HeadlessException;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPasswordField;
import javax.swing.JSeparator;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;

import com.google.gdata.util.AuthenticationException;
import com.google.gdata.util.ServiceException;

import net.miginfocom.swing.MigLayout;


/**
 * Window to collect Google Login and Password. 
 *  
 * @author Chris Martin
 */

public class GLogin extends JDialog implements ActionListener   //WindowListener, MouseListener, KeyListener
{
   
    private static final long serialVersionUID = 3L;
        
    private JTextField     username;
    private JPasswordField password;
    private JButton        login;

    private GCalendars	gCalendars;
    private boolean	validated = false;
    
    
    // Constructors
    
    public GLogin(JFrame parent, String glogon)
    {
	super(parent, true);
	
        Container container = getContentPane();
        container.setLayout(new MigLayout("", "[][:270:][::135]","[][][][]"));
    
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
        setTitle(Gcal2xls.res.getString("login.title"));
        setModal(true);
        setDefaultCloseOperation(JDialog.HIDE_ON_CLOSE);
    
        // Login detail separator
        container.add(new JHeading(Gcal2xls.res.getString("login.heading")), "split, span, gaptop 10");
        container.add(new JSeparator(), "growx, wrap, gaptop 10");
    
        // username
        username = new JTextField();
        username.setText(glogon);
        username.setToolTipText(Gcal2xls.res.getString("login.id.tooltip"));
        container.add(new JLabel(Gcal2xls.res.getString("login.id.label")) );
        container.add(username, "growx");
        container.add(new JHint(Gcal2xls.res.getString("login.id.hint")), "wrap");
    
        // password
      	password = new JPasswordField();
        password.setToolTipText(Gcal2xls.res.getString("login.passwd.tooltip"));
        container.add(new JLabel(Gcal2xls.res.getString("login.passwd.label")));
        container.add(password, "growx");
        container.add(new JHint(Gcal2xls.res.getString("login.passwd.hint")), "wrap");
        
        //The Login Button
        login = new JButton(Gcal2xls.res.getString("login.button"));
        login.addActionListener(this);
        container.add(login, "gaptop 10, skip 1, align right, gapbottom 15");

        pack();
        setResizable(false);
        
        transferFocus();
        
        
        // Set the focus to the appropriate field
        if (username.getText().length() == 0)
            // set the focus on the user name
            username.requestFocus();
        else
            // put the focus on the password
            password.requestFocus();
    }


    public void actionPerformed(ActionEvent event)
    {
	Object object = event.getSource();

	if ( object == login )
	{
            // Check password and name fields to see if we have any.
            if ( (username.getText().length() != 0) && (password.getPassword().length != 0) )
            {
                // make a busy cursor
                setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
                
                // invalidate the credentials
                validated = false;
                
                // fetch the list of gCalenders
                try
                {
                    
                    gCalendars = new GCalendars(Gcal2xls.googleApiId);
                    gCalendars.setIncludeHiddenCalendars(true);
                    gCalendars.update(username.getText(), new String(password.getPassword()));
                  
                    // assume that all returned OK
                    validated = true;
                    setVisible(false);
               }
               catch (AuthenticationException e)
               {
                   displayError(Gcal2xls.res.getString("error.auth"), Gcal2xls.res.getString("error.auth.invalid"));
                   return;
               }
               catch (IOException e)
               {
                   displayError(Gcal2xls.res.getString("error.io"), e.getLocalizedMessage());
                   return;
               }
               catch (ServiceException e)
               {
                   displayError(Gcal2xls.res.getString("error.service"), e.getLocalizedMessage());
                   return;
               }
               finally
               {
                   setCursor(Cursor.getDefaultCursor());
               }
            }
            else
            {
                displayError(Gcal2xls.res.getString("error.auth"), Gcal2xls.res.getString("error.auth.missing"));
                return;
            }
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
    
    
    
    // Getters
    
    public String getUsername()
    {
	return username.getText();
    }
    
    public char[] getPassword()
    {
	return password.getPassword();
    }
    
    public GCalendars getCalendars()
    {
	return gCalendars;
    }
    
    public boolean isValidated()
    {
	return validated;
    }
}
