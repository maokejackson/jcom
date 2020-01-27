package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookBodyFormat;
import com.dtxmaker.jcom.outlook.constant.OutlookMarkInterval;
import com.jacob.com.Dispatch;

import java.util.Date;

/**
 * Represents a post in a public folder that others may browse.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.postitem">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.postitem</a>
 */
public class OutlookPost extends AbstractOutlookInternalItem
{
    OutlookPost(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Clears the index of the conversation thread for the post.
     */
    public void clearConversationIndex()
    {
        call("ClearConversationIndex");
    }

    /**
     * Clears the item as a task.
     */
    public void clearTaskFlag()
    {
        call("ClearTaskFlag");
    }

    /**
     * Executes the Forward action for an item and returns the resulting copy.
     *
     * @return the new mail item.
     */
    public OutlookPost forward()
    {
        return new OutlookPost(callDispatch("Forward"));
    }

    /**
     * Marks it as a task and assigns a task interval for the object.
     *
     * @param markInterval the task interval
     */
    public void markAsTask(OutlookMarkInterval markInterval)
    {
        call("MarkAsTask", markInterval);
    }

    /**
     * Sends (posts) the PostItem object.
     */
    public void post()
    {
        call("Post");
    }

    /**
     * Creates a reply, pre-addressed to the original sender, from the original message.
     *
     * @return A MailItem object that represents the reply.
     */
    public OutlookMail reply()
    {
        return new OutlookMail(callDispatch("Reply"));
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    /**
     * Sets the format of the body text.
     *
     * @param bodyFormat the format of the body text
     */
    public void setBodyFormat(OutlookBodyFormat bodyFormat)
    {
        put("BodyFormat", bodyFormat);
    }

    /**
     * Returns the format of the body text.
     */
    public OutlookBodyFormat getBodyFormat()
    {
        return getConstant("BodyFormat", OutlookBodyFormat.class);
    }

    /**
     * Sets the date and time at which the item becomes invalid and can be deleted.
     *
     * @param expiryTime the date and time at which the item becomes invalid and can be deleted
     */
    public void setExpiryTime(Date expiryTime)
    {
        put("ExpiryTime", expiryTime);
    }

    /**
     * Returns the date and time at which the item becomes invalid and can be deleted.
     */
    public Date getExpiryTime()
    {
        return getDate("ExpiryTime");
    }

    /**
     * Sets the HTML body of the specified item.
     *
     * @param htmlBody the HTML body
     */
    public void setHtmlBody(String htmlBody)
    {
        put("HTMLBody", htmlBody);
    }

    /**
     * Returns the HTML body of the specified item.
     */
    public String getHtmlBody()
    {
        return getString("HTMLBody");
    }

    /**
     * Sets the Internet code page used by the item.
     *
     * @param internetCodePage the Internet code page
     */
    public void setInternetCodePage(long internetCodePage)
    {
        put("InternetCodepage", internetCodePage);
    }

    /**
     * Returns the Internet code page used by the item.
     */
    public long getInternetCodePage()
    {
        return getLong("InternetCodepage");
    }

    /**
     * Indicates whether the item is marked as a task.
     */
    public boolean isMarkedAsTask()
    {
        return getBoolean("IsMarkedAsTask");
    }

    /**
     * Returns the date and time at which the item was received.
     */
    public Date getReceivedTime()
    {
        return getDate("ReceivedTime");
    }

    /**
     * Sets if the reminder overrides the default reminder behavior for the item.
     *
     * @param reminderOverrideDefault <code>true</code> if the reminder overrides the default reminder.
     */
    public void setReminderOverrideDefault(boolean reminderOverrideDefault)
    {
        put("ReminderOverrideDefault", reminderOverrideDefault);
    }

    /**
     * Returns <code>true</code> if the reminder overrides the default reminder behavior for the item.
     */
    public boolean isReminderOverrideDefault()
    {
        return getBoolean("ReminderOverrideDefault");
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

    public void setReminderTime(Date reminderTime)
    {
        put("ReminderTime", reminderTime);
    }

    public Date getReminderTime()
    {
        return getDate("ReminderTime");
    }

    /**
     * Returns the email address of the sender of the Outlook item.
     *
     * @return the email address of the sender
     */
    public String getSenderEmailAddress()
    {
        return getString("SenderEmailAddress");
    }

    /**
     * Returns the type of entry for the email address of the sender of the Outlook item, such as 'SMTP' for Internet address, 'EX' for a Microsoft Exchange server address, etc.
     */
    public String getSenderEmailType()
    {
        return getString("SenderEmailType");
    }

    /**
     * Returns the display name of the sender for the Outlook item.
     */
    public String getSenderName()
    {
        return getString("SenderName");
    }

    /**
     * Returns the date and time on which the Outlook item was sent.
     */
    public Date getSentOn()
    {
        return getDate("SentOn");
    }

    /**
     * Sets the completion date of the task for this item.
     *
     * @param taskCompletedDate the completion date
     */
    public void setTaskCompletedDate(Date taskCompletedDate)
    {
        put("TaskCompletedDate", taskCompletedDate);
    }

    /**
     * Returns the completion date of the task for this item.
     */
    public Date getTaskCompletedDate()
    {
        return getDate("TaskCompletedDate");
    }

    /**
     * Sets the due date of the task for this item.
     *
     * @param taskDueDate the due date
     */
    public void setTaskDueDate(Date taskDueDate)
    {
        put("TaskDueDate", taskDueDate);
    }

    /**
     * Returns the due date of the task for this item.
     */
    public Date getTaskDueDate()
    {
        return getDate("TaskDueDate");
    }

    /**
     * Sets the start date of the task for this item.
     *
     * @param taskStartDate the start date
     */
    public void setTaskStartDate(Date taskStartDate)
    {
        put("TaskStartDate", taskStartDate);
    }

    /**
     * Returns the start date of the task for this item.
     */
    public Date getTaskStartDate()
    {
        return getDate("TaskStartDate");
    }

    /**
     * Sets the subject of the task for the item.
     *
     * @param taskSubject the subject
     */
    public void setTaskSubject(String taskSubject)
    {
        put("TaskSubject", taskSubject);
    }

    /**
     * Returns the subject of the task for the item.
     */
    public String getTaskSubject()
    {
        return getString("TaskSubject");
    }

    /**
     * Sets the ordinal value of the task for the item.
     *
     * @param toDoTaskOrdinal the ordinal value of the task
     */
    public void setToDoTaskOrdinal(Date toDoTaskOrdinal)
    {
        put("ToDoTaskOrdinal", toDoTaskOrdinal);
    }

    /**
     * Returns the ordinal value of the task for the item.
     */
    public Date getToDoTaskOrdinal()
    {
        return getDate("ToDoTaskOrdinal");
    }
}
