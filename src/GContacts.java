import java.io.IOException;
import java.net.URL;
import java.util.HashMap;

import com.google.gdata.client.contacts.ContactsService;
import com.google.gdata.data.contacts.ContactEntry;
import com.google.gdata.data.contacts.ContactFeed;
import com.google.gdata.data.extensions.Email;
import com.google.gdata.util.ServiceException;

/**
 * @author Chris Martin
 *
 */

public class GContacts
{
    private static final int MAX_ENTRIES = 100;
    private String googleApiId;

    private HashMap<String, String> contacts = new HashMap<String, String>(MAX_ENTRIES);
    
    public GContacts(String googleApiId)
    {
	this.googleApiId = googleApiId;
    }
    
    public void update(String username, String password) throws IOException, ServiceException
    {
	// Create a new Contacts service
	ContactsService contactService = new ContactsService(googleApiId);
	contactService.setReadTimeout(10000);	// allow 10 sec's to collect and read the data
	
	// authenticate
	contactService.setUserCredentials(username, password);
		  
	// Get a list of all entries
	URL contactsUrl = new URL("https://www.google.com/m8/feeds/contacts/" + username + "/base?max-results=" + MAX_ENTRIES);

	while (contactsUrl != null)
	{
	    ContactFeed resultFeed =  contactService.getFeed(contactsUrl, ContactFeed.class);
    
    	    // for each of the contacts returned
    	    for( ContactEntry entry :  resultFeed.getEntries() )
    	    {
    		String name = entry.getTitle().getPlainText().trim();
    	      
    		// there may be multiple address per person so we will add each email to the map
    		for ( Email email : entry.getEmailAddresses() )
    		{
    		    if ( name.isEmpty() )
    		    {
    			// no name - so use the email address
    			contacts.put(email.getAddress(), email.getAddress());
    		    }
    		    else
    		    {
    			// we have a valid name in contacts so use it
    			contacts.put(email.getAddress(), name);
    		    }
    		}
    	    }

    	    // see if there is more to fetch
    	    if (resultFeed.getNextLink() == null)
    		contactsUrl = null;
    	    else
    		contactsUrl = new URL(resultFeed.getNextLink().getHref());
	}
    }
    
    public String lookup (String email)
    {
	return contacts.get(email);
    }
}
