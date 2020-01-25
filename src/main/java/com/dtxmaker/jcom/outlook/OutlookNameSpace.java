package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookAutoDiscoverConnectionMode;
import com.dtxmaker.jcom.outlook.constant.OutlookDefaultFolderType;
import com.dtxmaker.jcom.outlook.constant.OutlookExchangeConnectionMode;
import com.dtxmaker.jcom.outlook.constant.OutlookStoreType;
import com.jacob.com.Dispatch;

/**
 * Represents an abstract root object for any data source.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.namespace">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.namespace</a>
 */
public class OutlookNameSpace extends Outlook
{
    OutlookNameSpace(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Adds a Personal Folders (.pst) file to the current profile.
     *
     * @param store The path of the .pst file to be added to the profile. If the .pst file does not exist, Microsoft Outlook creates it.
     */
    public void addStore(String store)
    {
        call("AddStore", store);
    }

    /**
     * Adds a Personal Folders file (.pst) in the specified format to the current profile.
     *
     * @param store The path of the .pst file to be added to the profile. If the .pst file does not exist, Microsoft Outlook creates it.
     * @param type  The format in which the data file should be created.
     */
    public void addStore(String store, OutlookStoreType type)
    {
        call("AddStoreEx", store, type.getValue());
    }

    /**
     * Returns a Boolean value that indicates if two entry ID values refer to the same Outlook item.
     *
     * @param firstEntryId  The first entry ID to be compared.
     * @param secondEntryId The second entry ID to be compared.
     * @return <code>true</code> if the entry ID values refer to the same Outlook item; otherwise, <code>false</code>.
     */
    public boolean compareEntryIDs(String firstEntryId, String secondEntryId)
    {
        return callBoolean("CompareEntryIDs", firstEntryId, secondEntryId);
    }

    /**
     * Creates a Recipient object.
     *
     * @param recipientName The name of the recipient; it can be a string representing the display name, the alias, or the full SMTP email address of the recipient.
     * @return A Recipient object that represents the new recipient.
     */
    public OutlookRecipient createRecipient(String recipientName)
    {
        return new OutlookRecipient(callDispatch("CreateRecipient", recipientName));
    }

    /**
     * Returns an AddressEntry object that represents the address entry for the specified ID.
     *
     * @param id Used to identify an address entry that is maintained for the session.
     * @return An AddressEntry that has the ID property that matches the specified ID.
     */
    public OutlookAddressEntry getAddressEntry(String id)
    {
        return new OutlookAddressEntry(callDispatch("GetAddressEntryFromID", id));
    }

    /**
     * Returns a Folder object that represents the default folder of the requested type for the current profile; for example, obtains the default Calendar folder for the user who is currently logged on.
     *
     * @param folderType The type of default folder to return.
     * @return A Folder object that represents the default folder of the requested type for the current profile.
     */
    public OutlookDefaultFolder getDefaultFolder(OutlookDefaultFolderType folderType)
    {
        return new OutlookDefaultFolder(callDispatch("GetDefaultFolder", folderType));
    }

    /**
     * Returns a Folder object identified by the specified entry ID (if valid).
     *
     * @param entryId The EntryID of the folder.
     * @return A Folder object that represents the specified folder.
     * @see OutlookFolder#getEntryId()
     */
    public OutlookFolder getFolder(String entryId)
    {
        return new OutlookFolder(callDispatch("GetFolderFromID", entryId));
    }

    /**
     * Returns a Folder object identified by the specified entry ID (if valid).
     *
     * @param entryId The EntryID of the folder.
     * @param storeId The StoreID for the folder.
     * @return A Folder object that represents the specified folder.
     * @see OutlookFolder#getEntryId()
     * @see OutlookFolder#getStoreId()
     */
    public OutlookFolder getFolder(String entryId, String storeId)
    {
        return new OutlookFolder(callDispatch("GetFolderFromID", entryId, storeId));
    }

    /**
     * Returns a Microsoft Outlook item identified by the specified entry ID (if valid).
     *
     * @param entryId The EntryID of the item.
     * @return An Object value that represents the specified Outlook item.
     * @see OutlookFolder#getEntryId()
     */
    public OutlookItem getItem(String entryId)
    {
        return new OutlookItem(callDispatch("GetItemFromID", entryId));
    }

    /**
     * Returns a Microsoft Outlook item identified by the specified entry ID (if valid).
     *
     * @param entryId The EntryID of the item.
     * @param storeId The StoreID for the folder. EntryIDStore usually must be provided when retrieving an item based on its MAPI IDs.
     * @return An Object value that represents the specified Outlook item.
     * @see OutlookFolder#getEntryId()
     * @see OutlookFolder#getStoreId()
     */
    public OutlookItem getItem(String entryId, String storeId)
    {
        return new OutlookItem(callDispatch("GetItemFromID", entryId, storeId));
    }

    /**
     * Returns the Recipient object that is identified by the specified entry ID (if valid).
     *
     * @param entryId The EntryID of the recipient.
     * @return A Recipient object that represents the specified recipient.
     * @see OutlookRecipient#getEntryId()
     */
    public OutlookRecipient getRecipient(String entryId)
    {
        return new OutlookRecipient(callDispatch("GetRecipientFromID", entryId));
    }

    /**
     * Returns a Folder object that represents the specified default folder for the specified user.
     *
     * @param recipient  The owner of the folder. Note that the Recipient object must be resolved.
     * @param folderType The type of folder.
     * @return A Folder object that represents the specified default folder for the specified user.
     */
    public OutlookDefaultFolder getSharedDefaultFolder(OutlookRecipient recipient, OutlookDefaultFolderType folderType)
    {
        return new OutlookDefaultFolder(callDispatch("GetSharedDefaultFolder", recipient, folderType));
    }

    /**
     * Logs the user off from the current MAPI session.
     */
    public void logoff()
    {
        call("Logoff");
    }

    /**
     * Initiates immediate delivery of all undelivered messages submitted in the current session, and immediate receipt of mail for all accounts in the current profile.
     *
     * @param showProgressDialog Indicates whether the Outlook Send/Receive Progress dialog box should be displayed, regardless of user settings.
     */
    public void sendAndReceive(boolean showProgressDialog)
    {
        call("SendAndReceive", showProgressDialog);
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    /**
     * Returns an enumeration that specifies the type of connection for auto-discovery of the Microsoft Exchange server that hosts the primary Exchange account.
     */
    public OutlookAutoDiscoverConnectionMode getAutoDiscoverConnectionMode()
    {
        return getConstant("AutoDiscoverConnectionMode", OutlookAutoDiscoverConnectionMode.class);
    }

    /**
     * Returns a String that represents information in XML retrieved from the auto-discovery service for the Microsoft Exchange server that hosts the primary Exchange account.
     */
    public String getAutoDiscoverXml()
    {
        return getString("AutoDiscoverXml");
    }

    /**
     * Sets a Categories object that represents the set of Category objects that are available to the namespace.
     *
     * @param categories the set of Category objects.
     */
    public void setCategories(OutlookCategories categories)
    {
        put("Categories", categories);
    }

    /**
     * Returns a Categories object that represents the set of Category objects that are available to the namespace.
     */
    public OutlookCategories getCategories()
    {
        return new OutlookCategories(getDispatch("Categories"));
    }

    /**
     * Returns a String representing the name of the current profile.
     */
    public String getCurrentProfileName()
    {
        return getString("CurrentProfileName");
    }

    /**
     * Returns the display name of the currently logged-on user as a Recipient object.
     */
    public OutlookRecipient getCurrentUser()
    {
        return new OutlookRecipient(getDispatch("CurrentUser"));
    }

    /**
     * Returns an enumeration that indicates the connection mode of the user's primary Exchange account.
     */
    public OutlookExchangeConnectionMode getExchangeConnectionMode()
    {
        return getConstant("ExchangeConnectionMode", OutlookExchangeConnectionMode.class);
    }

    /**
     * Returns a String value that represents the name of the Exchange server that hosts the primary Exchange account mailbox.
     */
    public String getExchangeMailboxServerName()
    {
        return getString("ExchangeMailboxServerName");
    }

    /**
     * Returns a String value that represents the full version number of the Exchange server that hosts the primary Exchange account mailbox.
     */
    public String getExchangeMailboxServerVersion()
    {
        return getString("ExchangeMailboxServerVersion");
    }

    /**
     * Returns the Folders collection that represents all the folders contained in the specified NameSpace.
     */
    public OutlookFolders getFolders()
    {
        return new OutlookFolders(getDispatch("Folders"));
    }

    /**
     * Returns a Boolean indicating <code>true</code> if Outlook is offline (not connected to an Exchange server), and <code>false</code> if online (connected to an Exchange server).
     */
    public boolean isOffline()
    {
        return getBoolean("Offline");
    }

    /**
     * Returns a String indicating the type of the specified object.
     */
    public String getType()
    {
        return getString("Type");
    }
}
