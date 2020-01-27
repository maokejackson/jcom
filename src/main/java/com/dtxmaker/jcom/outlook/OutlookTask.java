package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookTaskDelegationState;
import com.dtxmaker.jcom.outlook.constant.OutlookTaskOwnership;
import com.dtxmaker.jcom.outlook.constant.OutlookTaskResponse;
import com.dtxmaker.jcom.outlook.constant.OutlookTaskStatus;
import com.jacob.com.Dispatch;

import java.util.Date;

/**
 * Represents a task (an assigned, delegated, or self-imposed task to be performed within a specified time frame) in a Tasks folder.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.taskitem">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.taskitem</a>
 */
public class OutlookTask extends AbstractOutlookInternalItem
{
    OutlookTask(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    /**
     * sets the actual effort spent on the task.
     *
     * @param actualWork the actual effort spent
     */
    public void setActualWork(int actualWork)
    {
        put("ActualWork", actualWork);
    }

    /**
     * Returns the actual effort spent on the task.
     */
    public int getActualWork()
    {
        return getInt("ActualWork");
    }

    /**
     * Sets the text of the card data for the task.
     *
     * @param cardData the text of the card data
     */
    public void setCardData(String cardData)
    {
        put("CardData", cardData);
    }

    /**
     * Returns the text of the card data for the task.
     */
    public String getCardData()
    {
        return getString("CardData");
    }

    /**
     * Sets whether the task is completed.
     *
     * @param complete <code>true</code> to indicate the task is completed
     */
    public void setComplete(boolean complete)
    {
        put("Complete", complete);
    }

    /**
     * Indicates whether the task is completed.
     */
    public boolean isComplete()
    {
        return getBoolean("Complete");
    }

    /**
     * Sets the contact names associated with the Outlook item.
     *
     * @param contactNames the contact names
     */
    public void setContactNames(String contactNames)
    {
        put("ContactNames", contactNames);
    }

    /**
     * Returns the contact names associated with the Outlook item.
     */
    public String getContactNames()
    {
        return getString("ContactNames");
    }

    /**
     * Sets the completion date of the task.
     *
     * @param dateCompleted the completion date
     */
    public void setDateCompleted(Date dateCompleted)
    {
        put("DateCompleted", dateCompleted);
    }

    /**
     * Returns the completion date of the task.
     */
    public Date getDateCompleted()
    {
        return getDate("DateCompleted");
    }

    /**
     * Returns the delegation state of the task.
     */
    public OutlookTaskDelegationState getDelegationState()
    {
        return getConstant("DelegationState", OutlookTaskDelegationState.class);
    }

    /**
     * Returns the display name of the delegator for the task.
     */
    public String getDelegator()
    {
        return getString("Delegator");
    }

    /**
     * Sets the due date for the task.
     *
     * @param dueDate the due date
     */
    public void setDueDate(Date dueDate)
    {
        put("DueDate", dueDate);
    }

    /**
     * Returns the due date for the task.
     */
    public Date getDueDate()
    {
        return getDate("DueDate");
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
     * Indicates if the task is a recurring task.
     */
    public boolean isRecurring()
    {
        return getBoolean("IsRecurring");
    }

    /**
     * Sets the position in the view (ordinal) for the task.
     *
     * @param ordinal the position in the view
     */
    public void setOrdinal(int ordinal)
    {
        put("Ordinal", ordinal);
    }

    /**
     * Returns the position in the view (ordinal) for the task.
     */
    public int getOrdinal()
    {
        return getInt("Ordinal");
    }

    /**
     * Sets the owner for the task.
     *
     * @param owner the owner
     */
    public void setOwner(String owner)
    {
        put("Owner", owner);
    }

    /**
     * Returns the owner for the task.
     */
    public String getOwner()
    {
        return getString("Owner");
    }

    /**
     * Returns the ownership state of the task.
     */
    public OutlookTaskOwnership getOwnership()
    {
        return getConstant("Ownership", OutlookTaskOwnership.class);
    }

    /**
     * Sets the percentage of the task completed at the current date and time.
     *
     * @param percentComplete the percentage of the task completed
     */
    public void setPercentComplete(int percentComplete)
    {
        put("PercentComplete", percentComplete);
    }

    /**
     * Returns the percentage of the task completed at the current date and time.
     */
    public int getPercentComplete()
    {
        return getInt("PercentComplete");
    }

    /**
     * Returns all the recipients for the Outlook item.
     */
    public final OutlookRecipients getRecipients()
    {
        return new OutlookRecipients(callDispatch("Recipients"));
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
     * Returns the overall status of the response to the specified task request.
     */
    public OutlookTaskResponse getResponseState()
    {
        return getConstant("ResponseState", OutlookTaskResponse.class);
    }

    /**
     * Sets the free-form text string associating the owner of a task with a role for the task.
     *
     * @param role the owner of a task with a role
     */
    public void setRole(String role)
    {
        put("Role", role);
    }

    /**
     * Returns the free-form text string associating the owner of a task with a role for the task.
     */
    public String getRole()
    {
        return getString("Role");
    }

    /**
     * Sets the Microsoft Schedule+ priority for the task.
     *
     * @param schedulePlusPriority the Microsoft Schedule+ priority
     */
    public void setSchedulePlusPriority(String schedulePlusPriority)
    {
        put("SchedulePlusPriority", schedulePlusPriority);
    }

    /**
     * Returns the Microsoft Schedule+ priority for the task.
     */
    public String getSchedulePlusPriority()
    {
        return getString("SchedulePlusPriority");
    }

    /**
     * Sets the start date for the task.
     *
     * @param startDate the start date
     */
    public void setStartDate(Date startDate)
    {
        put("StartDate", startDate);
    }

    /**
     * Returns the start date for the task.
     */
    public Date getStartDate()
    {
        return getDate("StartDate");
    }

    /**
     * Sets the status for the task.
     *
     * @param status the status
     */
    public void setStatus(OutlookTaskStatus status)
    {
        put("Status", status);
    }

    /**
     * Returns the status for the task.
     */
    public OutlookTaskStatus getStatus()
    {
        return getConstant("Status", OutlookTaskStatus.class);
    }

    /**
     * Sets display names for recipients who will receive status upon completion of the task.
     *
     * @param statusOnCompletionRecipients display names (semicolon-delimited)
     */
    public void setStatusOnCompletionRecipients(String statusOnCompletionRecipients)
    {
        put("StatusOnCompletionRecipients", statusOnCompletionRecipients);
    }

    /**
     * Returns display names (semicolon-delimited) for recipients who will receive status upon completion of the task.
     */
    public String getStatusOnCompletionRecipients()
    {
        return getString("StatusOnCompletionRecipients");
    }

    /**
     * Sets display names for recipients who receive status updates for the task.
     *
     * @param statusUpdateRecipients display names (semicolon-delimited)
     */
    public void setStatusUpdateRecipients(String statusUpdateRecipients)
    {
        put("StatusUpdateRecipients", statusUpdateRecipients);
    }

    /**
     * Returns display names (semicolon-delimited) for recipients who receive status updates for the task.
     */
    public String getStatusUpdateRecipients()
    {
        return getString("StatusUpdateRecipients");
    }

    /**
     * Sets whether the task is a team task.
     *
     * @param teamTask <code>true</code> if the task is a team task
     */
    public void setTeamTask(boolean teamTask)
    {
        put("TeamTask", teamTask);
    }

    /**
     * Indicates if the task is a team task.
     */
    public boolean isTeamTask()
    {
        return getBoolean("TeamTask");
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

    /**
     * Sets the total work for the task.
     *
     * @param totalWork the total work
     */
    public void setTotalWork(int totalWork)
    {
        put("TotalWork", totalWork);
    }

    /**
     * Returns the total work for the task.
     */
    public int getTotalWork()
    {
        return getInt("TotalWork");
    }
}
