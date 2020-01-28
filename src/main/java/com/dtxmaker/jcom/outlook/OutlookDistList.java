package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookMarkInterval;
import com.jacob.com.Dispatch;

import java.util.Date;

/**
 * Represents a distribution list in a Contacts folder.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.distlistitem">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.distlistitem</a>
 */
public class OutlookDistList extends AbstractOutlookInternalItem
{
    OutlookDistList(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Adds a new member to the specified distribution list.
     *
     * @param recipient the recipient to be added to the list
     */
    public void addMember(OutlookRecipient recipient)
    {
        call("AddMember", recipient);
    }

    /**
     * Adds new members to a distribution list.
     *
     * @param recipients the members to be added to the distribution list
     */
    public void addMembers(OutlookRecipients recipients)
    {
        call("AddMembers", recipients);
    }

    /**
     * Returns a member in a distribution list.
     *
     * @param index he index number of the member to be retrieved
     * @return the specified member.
     */
    public OutlookRecipient getMember(int index)
    {
        return new OutlookRecipient(callDispatch("GetMember", index));
    }

    /**
     * Clears the item as a task.
     */
    public void clearTaskFlag()
    {
        call("ClearTaskFlag");
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
     * Removes an individual member from a given distribution list.
     *
     * @param recipient the Recipient to be removed from the distribution list
     */
    public void removeMember(OutlookRecipient recipient)
    {
        call("RemoveMember", recipient);
    }

    /**
     * Removes members from a distribution list.
     *
     * @param recipients the members to be removed from the distribution list
     */
    public void removeMembers(OutlookRecipients recipients)
    {
        call("RemoveMembers", recipients);
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    /**
     * Sets the display name of a distribution list.
     *
     * @param name the display name
     */
    public void setName(String name)
    {
        put("DLName", name);
    }

    /**
     * Returns the display name of a distribution list.
     */
    public String getName()
    {
        return getString("DLName");
    }

    /**
     * Indicates whether the item is marked as a task.
     */
    public boolean isMarkedAsTask()
    {
        return getBoolean("IsMarkedAsTask");
    }

    /**
     * Returns the number of members in a distribution list.
     */
    public int getMemberCount()
    {
        return getInt("MemberCount");
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
