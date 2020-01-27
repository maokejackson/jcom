package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookDownloadState;
import com.dtxmaker.jcom.outlook.constant.OutlookInspectorClose;
import com.dtxmaker.jcom.outlook.constant.OutlookRemoteStatus;
import com.dtxmaker.jcom.outlook.constant.OutlookSaveAsType;
import com.jacob.com.Dispatch;

import java.util.Date;

public class OutlookItem extends Outlook
{
    OutlookItem(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Closes and optionally saves changes to the Outlook item.
     *
     * @param close The close behavior.
     */
    public final void close(OutlookInspectorClose close)
    {
        call("Close", close);
    }

    /**
     * Creates another instance of an object.
     */
    public final OutlookItem copy()
    {
        return new OutlookItem(callDispatch("Copy"));
    }

    /**
     * Removes the item from the folder that contains the item.
     */
    public final void delete()
    {
        call("Delete");
    }

    /**
     * Displays a new Inspector object for the item.
     */
    public final void display()
    {
        call("Display");
    }

    /**
     * Displays a new Inspector object for the item.
     *
     * @param modal <code>true</code> to make the window modal.
     */
    public final void display(boolean modal)
    {
        call("Display", modal);
    }

    /**
     * Moves a Microsoft Outlook item to a new <code>folder</code>.
     *
     * @param folder the destination folder
     */
    public final void move(OutlookFolder folder)
    {
        call("Move", folder);
    }

    /**
     * Prints the Outlook item using all default settings.
     */
    public final void printOut()
    {
        call(" PrintOut");
    }

    /**
     * Saves the Microsoft Outlook item to the current folder or, if this is a new item, to the Outlook default folder for the item type.
     */
    public final void save()
    {
        call("Save");
    }

    /**
     * Saves the Microsoft Outlook item to the specified path and in the format of MSG (.msg).
     */
    public final void saveAs(String path)
    {
        call("SaveAs", path);
    }

    /**
     * Saves the Microsoft Outlook item to the specified path and in the format of the specified file type.
     */
    public final void saveAs(String path, OutlookSaveAsType type)
    {
        call("SaveAs", path, type);
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    /**
     * Determines if the item is a winner of an automatic conflict resolution.
     */
    public final boolean isAutoResolvedWinner()
    {
        return getBoolean("AutoResolvedWinner");
    }

    /**
     * Sets the clear-text body of the Outlook item.
     *
     * @param body The clear-text body of the Outlook item.
     */
    public final void setBody(String body)
    {
        put("Body", body);
    }

    /**
     * Returns the clear-text body of the Outlook item.
     */
    public final String getBody()
    {
        return getString("Body");
    }

    /**
     * Sets the categories assigned to the Outlook item.
     *
     * @param categories The categories assigned to the Outlook item.
     */
    public final void setCategories(String categories)
    {
        put("Categories", categories);
    }

    /**
     * Returns the categories assigned to the Outlook item.
     */
    public final String getCategories()
    {
        return getString("Categories");
    }

    /**
     * Returns the creation time for the Outlook item.
     */
    public final Date getCreationTime()
    {
        return getDate("CreationTime");
    }

    /**
     * Returns the download state of the item.
     */
    public final OutlookDownloadState getDownloadState()
    {
        return getConstant("DownloadState", OutlookDownloadState.class);
    }

    /**
     * Returns the unique Entry ID of the object.
     */
    public final String getEntryId()
    {
        return getString("EntryID");
    }

    /**
     * Determines if the item is in conflict.
     */
    public final boolean isConflict()
    {
        return getBoolean("IsConflict");
    }

    /**
     * Returns the date and time that the Outlook item was last modified.
     */
    public final Date getLastModificationTime()
    {
        return getDate("LastModificationTime");
    }

    /**
     * Sets the status of an item once it is received by a remote user.
     *
     * @param markForDownload The status of an item.
     */
    public final void setMarkForDownload(OutlookRemoteStatus markForDownload)
    {
        put("MarkForDownload", markForDownload);
    }

    /**
     * Returns the status of an item once it is received by a remote user.
     */
    public final OutlookRemoteStatus getMarkForDownload()
    {
        return getConstant("MarkForDownload", OutlookRemoteStatus.class);
    }

    /**
     * Sets the message class for the Outlook item.
     *
     * @param messageClass The message class.
     */
    public final void setMessageClass(String messageClass)
    {
        put("MessageClass", messageClass);
    }

    /**
     * Returns the message class for the Outlook item.
     */
    public final String getMessageClass()
    {
        return getString("MessageClass");
    }

    /**
     * Returns <code>true</code> if the Outlook item has not been modified since the last save.
     */
    public final boolean isSaved()
    {
        return getBoolean("Saved");
    }

    /**
     * Returns the size (in bytes) of the Outlook item.
     */
    public final long getSize()
    {
        return getLong("Size");
    }

    /**
     * Sets the subject for the Outlook item.
     *
     * @param subject the subject
     */
    public final void setSubject(String subject)
    {
        put("Subject", subject);
    }

    /**
     * Returns the subject for the Outlook item.
     */
    public final String getSubject()
    {
        return getString("Subject");
    }
}
