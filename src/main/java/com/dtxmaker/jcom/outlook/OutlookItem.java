package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.*;
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

    /**
     * Displays the Show Categories dialog box, which allows you to select categories that correspond to the subject of the item.
     */
    public final void showCategoriesDialog()
    {
        call("ShowCategoriesDialog");
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    /**
     * Returns all the attachments for the specified item.
     */
    public final OutlookAttachments getAttachments()
    {
        return new OutlookAttachments(getDispatch("Attachments"));
    }

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
     * Returns a String that uniquely identifies a Conversation object that the Outlook item belongs to.
     */
    public final String getConversationId()
    {
        return getString("ConversationID");
    }

    /**
     * Returns the relative position of the item within the conversation thread.
     */
    public final String getConversationIndex()
    {
        return getString("ConversationIndex");
    }

    /**
     * Returns the topic of the conversation thread of the Outlook item.
     */
    public final String getConversationTopic()
    {
        return getString("ConversationTopic");
    }

    /**
     * Returns the creation time for the Outlook item.
     */
    public final Date getCreationTime()
    {
        return getDate("CreationTime");
    }

    /**
     * Returns the unique Entry ID of the object.
     */
    public final String getEntryId()
    {
        return getString("EntryID");
    }

    /**
     * Sets the relative importance level for the Outlook item.
     *
     * @param importance The relative importance level for the Outlook item.
     */
    public final void setImportance(OutlookImportance importance)
    {
        put("Importance", importance);
    }

    /**
     * Returns the relative importance level for the Outlook item.
     */
    public final OutlookImportance getImportance()
    {
        return getConstant("Importance", OutlookImportance.class);
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
     * Sets the mileage for an item.
     *
     * @param mileage The mileage.
     */
    public final void setMileage(String mileage)
    {
        put("Mileage", mileage);
    }

    /**
     * Returns the mileage for an item.
     */
    public final String getMileage()
    {
        return getString("Mileage");
    }

    /**
     * Sets <code>true</code> to not age the Outlook item.
     *
     * @param noAging <code>true</code> to not age the Outlook item; Otherwise, <code>false</code>.
     */
    public final void setNoAging(boolean noAging)
    {
        put("NoAging", noAging);
    }

    /**
     * Returns <code>true</code> to not age the Outlook item.
     */
    public final boolean isNoAging()
    {
        return getBoolean("NoAging");
    }

    /**
     * Returns the build number of the Outlook application for an Outlook item.
     */
    public final long getOutlookInternalVersion()
    {
        return getLong("OutlookInternalVersion");
    }

    /**
     * Returns the major and minor version number of the Outlook application for an Outlook item.
     */
    public final String getOutlookVersion()
    {
        return getString("OutlookVersion");
    }

    /**
     * Sets if the reminder overrides the default reminder behavior for the item.
     *
     * @param reminderOverrideDefault <code>true</code> if the reminder overrides the default reminder.
     */
    public final void setReminderOverrideDefault(String reminderOverrideDefault)
    {
        put("ReminderOverrideDefault", reminderOverrideDefault);
    }

    /**
     * Returns <code>true</code> if the reminder overrides the default reminder behavior for the item.
     */
    public final String getReminderOverrideDefault()
    {
        return getString("ReminderOverrideDefault");
    }

    /**
     * Sets if the reminder should play a sound when it occurs for this item.
     *
     * @param reminderPlaySound <code>true</code> if the reminder should play a sound.
     */
    public final void setReminderPlaySound(boolean reminderPlaySound)
    {
        put("ReminderPlaySound", reminderPlaySound);
    }

    /**
     * Returns <code>true</code> if the reminder should play a sound when it occurs for this item.
     */
    public final boolean isReminderPlaySound()
    {
        return getBoolean("ReminderPlaySound");
    }

    /**
     * Sets if a reminder has been set for this item.
     *
     * @param reminderSet <code>true</code> if a reminder has been set.
     */
    public final void setReminderSet(boolean reminderSet)
    {
        put("ReminderSet", reminderSet);
    }

    /**
     * Returns <code>true</code> if a reminder has been set for this item.
     */
    public final boolean getReminderSet()
    {
        return getBoolean("ReminderSet");
    }

    /**
     * Sets the path and file name of the sound file to play when the reminder occurs for the Outlook item.
     *
     * @param reminderSoundFile the path and file name of the sound file.
     */
    public final void setReminderSoundFile(String reminderSoundFile)
    {
        put("ReminderSoundFile", reminderSoundFile);
    }

    /**
     * Returns the path and file name of the sound file to play when the reminder occurs for the Outlook item.
     */
    public final String getReminderSoundFile()
    {
        return getString("ReminderSoundFile");
    }

    /**
     * Returns <code>true</code> if the Outlook item has not been modified since the last save.
     */
    public final boolean isSaved()
    {
        return getBoolean("Saved");
    }

    /**
     * Sets the sensitivity for the Outlook item.
     *
     * @param sensitivity the sensitivity
     */
    public final void setSensitivity(OutlookSensitivity sensitivity)
    {
        put("Sensitivity", sensitivity);
    }

    /**
     * Returns the sensitivity for the Outlook item.
     */
    public final OutlookSensitivity getSensitivity()
    {
        return getConstant("Sensitivity", OutlookSensitivity.class);
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

    /**
     * Sets if the Outlook item has not been opened (read).
     *
     * @param unRead <code>true</code> if the Outlook item has not been opened (read)
     */
    public final void setUnRead(boolean unRead)
    {
        put("UnRead", unRead);
    }

    /**
     * Returns <code>true</code> if the Outlook item has not been opened (read).
     */
    public final boolean isUnRead()
    {
        return getBoolean("UnRead");
    }
}
