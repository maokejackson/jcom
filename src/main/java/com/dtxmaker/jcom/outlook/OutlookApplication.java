package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.library.LanguageSettings;
import com.dtxmaker.jcom.outlook.constant.OutlookItemType;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComFailException;
import com.jacob.com.Dispatch;

import static com.dtxmaker.jcom.outlook.constant.OutlookItemType.*;

/**
 * Represents the entire Microsoft Outlook application.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.application">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.application</a>
 */
public class OutlookApplication extends Outlook
{
    private static final String OUTLOOK_APPLICATION = "Outlook.Application";

    public static boolean isInstalled()
    {
        try
        {
            new ActiveXComponent(OUTLOOK_APPLICATION);
            return true;
        }
        catch (ComFailException e)
        {
            return false;
        }
    }

    public OutlookApplication()
    {
        super(new ActiveXComponent(OUTLOOK_APPLICATION));
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Copies a file from a specified location into a Microsoft Outlook store.
     *
     * @param filePath       The path name of the object you want to copy.
     * @param destFolderPath The location you want to copy the file to.
     */
    public void copyFile(String filePath, String destFolderPath)
    {
        call("CopyFile", filePath, destFolderPath);
    }

    /**
     * Creates and returns a new Microsoft Outlook item.
     *
     * @param itemType The Outlook item type for the new item.
     * @return An Object value that represents the new Outlook item.
     */
    private Dispatch createItem(OutlookItemType itemType)
    {
        return callDispatch("CreateItem", itemType);
    }

    /**
     * Creates and returns a new Microsoft Mail item.
     */
    public OutlookMail createMail()
    {
        return new OutlookMail(createItem(MAIL));
    }

    /**
     * Creates and returns a new Microsoft Contact item.
     */
    public OutlookContact createContact()
    {
        return new OutlookContact(createItem(CONTACT));
    }

    /**
     * Creates and returns a new Microsoft Appointment item.
     */
    public OutlookAppointment createAppointment()
    {
        return new OutlookAppointment(createItem(APPOINTMENT));
    }

    /**
     * Creates and returns a new Microsoft Task item.
     */
    public OutlookTask createTask()
    {
        return new OutlookTask(createItem(TASK));
    }

    /**
     * Creates and returns a new Microsoft Journal item.
     */
    public OutlookJournal createJournal()
    {
        return new OutlookJournal(createItem(JOURNAL));
    }

    /**
     * Creates and returns a new Microsoft Post item.
     */
    public OutlookPost createPost()
    {
        return new OutlookPost(createItem(POST));
    }

    /**
     * Creates and returns a new Microsoft Note item.
     */
    public OutlookNote createNote()
    {
        return new OutlookNote(createItem(NOTE));
    }

    /**
     * Returns a NameSpace object of the specified type.
     *
     * @return A NameSpace object that represents the specified namespace.
     */
    public OutlookNameSpace getNamespace()
    {
        return getSession();
    }

    /**
     * Returns a Boolean indicating if a search will be synchronous or asynchronous.
     *
     * @param lookInFolders The path name of the folders that the search will search through. You must enclose the folder path with single quotes.
     * @return <code>true</cdde> if the search is synchronous; otherwise, <code>false</code>.
     */
    public boolean isSearchSynchronous(String lookInFolders)
    {
        return callBoolean("IsSearchSynchronous", lookInFolders);
    }

    /**
     * Closes all currently open windows.
     */
    public void quit()
    {
        call("Quit");
    }

    /**
     * Refreshes the cache by obtaining the current definition from the Windows registry for one or all of the form regions that are defined for the local machine and the current user.
     *
     * @param regionName The internal name of the form region whose definition you want to refresh in the cache. To refresh all form region definitions, specify an empty string.
     */
    public void refreshFormRegionDefinition(String regionName)
    {
        call("RefreshFormRegionDefinition", regionName);
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    /**
     * Returns a String representing the name of the default profile name.
     */
    public String getDefaultProfileName()
    {
        return getString("DefaultProfileName");
    }

    /**
     * Returns a Boolean to indicate if an add-in or external caller is considered trusted by Outlook.
     */
    public boolean isTrusted()
    {
        return getBoolean("IsTrusted");
    }

    /**
     * Returns a LanguageSettings object for the application that contains the language-specific attributes of Outlook.
     */
    public LanguageSettings getLanguageSettings()
    {
        return new LanguageSettings(getDispatch("LanguageSettings"));
    }

    /**
     * Returns a String value that represents the display name for the object.
     */
    public String getName()
    {
        return getString("Name");
    }

    /**
     * Returns a String specifying the Microsoft Outlook globally unique identifier (GUID).
     */
    public String getProductCode()
    {
        return getString("ProductCode");
    }

    /**
     * Returns a TimeZones collection that represents the set of time zones supported by Outlook.
     */
    public OutlookTimeZones getTimeZones()
    {
        return new OutlookTimeZones(getDispatch("TimeZones"));
    }

    /**
     * Returns or sets a String indicating the number of the version.
     */
    public String getVersion()
    {
        return getString("Version");
    }
}
