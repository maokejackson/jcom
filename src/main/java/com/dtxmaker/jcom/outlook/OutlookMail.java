package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.*;
import com.jacob.com.Dispatch;

import java.util.Date;

/**
 * Represents a mail message.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem</a>
 */
public class OutlookMail extends AbstractOutlookInternalItem
{
    OutlookMail(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Clears the index of the conversation thread for the mail message.
     */
    public void clearConversationIndex()
    {
        call("ClearConversationIndex");
    }

    /**
     * Clears the MailItem object as a task.
     */
    public void clearTaskFlag()
    {
        call("ClearTaskFlag");
    }

    /**
     * Executes the Forward action for an item and returns the resulting copy as a MailItem object.
     *
     * @return A MailItem object that represents the new mail item.
     */
    public OutlookMail forward()
    {
        return new OutlookMail(callDispatch("Forward"));
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
     * Creates a reply, pre-addressed to the original sender, from the original message.
     *
     * @return A MailItem object that represents the reply.
     */
    public OutlookMail reply()
    {
        return new OutlookMail(callDispatch("Reply"));
    }

    /**
     * Creates a reply to all original recipients from the original message.
     *
     * @return A MailItem object that represents the reply.
     */
    public OutlookMail replyAll()
    {
        return new OutlookMail(callDispatch("ReplyAll"));
    }

    /**
     * Sends the email message.
     */
    public void send()
    {
        call("Send");
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    public void setAlternateRecipientAllowed(boolean alternateRecipientAllowed)
    {
        put("AlternateRecipientAllowed", alternateRecipientAllowed);
    }

    public boolean isAlternateRecipientAllowed()
    {
        return getBoolean("AlternateRecipientAllowed");
    }

    /**
     * Sets if the item was automatically forwarded.
     *
     * @param autoForwarded <code>true</code> if the item was automatically forwarded
     */
    public void setAutoForwarded(boolean autoForwarded)
    {
        put("AutoForwarded", autoForwarded);
    }

    /**
     * Returns <code>true</code> if the item was automatically forwarded.
     */
    public boolean isAutoForwarded()
    {
        return getBoolean("AutoForwarded");
    }

    public void setBcc(String bcc)
    {
        put("BCC", bcc);
    }

    public String getBcc()
    {
        return getString("BCC");
    }

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

    public void setCc(String cc)
    {
        put("CC", cc);
    }

    public String getCc()
    {
        return getString("CC");
    }

    /**
     * Sets the date and time the mail message is to be delivered.
     *
     * @param deferredDeliveryTime the date and time the mail message is to be delivered
     */
    public void setDeferredDeliveryTime(Date deferredDeliveryTime)
    {
        put("DeferredDeliveryTime", deferredDeliveryTime);
    }

    /**
     * Returns the date and time the mail message is to be delivered.
     */
    public Date getDeferredDeliveryTime()
    {
        return getDate("DeferredDeliveryTime");
    }

    /**
     * Sets if a copy of the mail message should be saved in Sent Items folder upon being sent.
     *
     * @param deleteAfterSubmit <code>true</code> if a copy of the mail message is not saved upon being sent, and <code>false</code> if a copy is saved in Sent Items folder
     */
    public void setDeleteAfterSubmit(boolean deleteAfterSubmit)
    {
        put("DeleteAfterSubmit", deleteAfterSubmit);
    }

    /**
     * Returns <code>true</code> if a copy of the mail message is not saved upon being sent, and <code>false</code> if a copy is saved in Sent Items folder.
     */
    public boolean isDeleteAfterSubmit()
    {
        return getBoolean("DeleteAfterSubmit");
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
     * Sets the requested action for a mail item.
     *
     * @param flagRequest the requested action
     */
    public void setFlagRequest(String flagRequest)
    {
        put("FlagRequest", flagRequest);
    }

    /**
     * Returns the requested action for a mail item.
     */
    public String getFlagRequest()
    {
        return getString("FlagRequest");
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
     * Sets whether the originator of the meeting item or mail message will receive a delivery report.
     *
     * @param originatorDeliveryReportRequested <code>true</code> to receive a delivery report
     */
    public void setOriginatorDeliveryReportRequested(boolean originatorDeliveryReportRequested)
    {
        put("OriginatorDeliveryReportRequested", originatorDeliveryReportRequested);
    }

    /**
     * Determines whether the originator of the meeting item or mail message will receive a delivery report.
     */
    public boolean isOriginatorDeliveryReportRequested()
    {
        return getBoolean("OriginatorDeliveryReportRequested");
    }

    public void setPermission(OutlookPermission permission)
    {
        put("Permission", permission);
    }

    public OutlookPermission getPermission()
    {
        return getConstant("Permission", OutlookPermission.class);
    }

    public void setPermissionService(OutlookPermissionService permissionService)
    {
        put("PermissionService", permissionService);
    }

    public OutlookPermissionService getPermissionService()
    {
        return getConstant("PermissionService", OutlookPermissionService.class);
    }

    public void setPermissionTemplateGuid(String permissionTemplateGuid)
    {
        put("PermissionTemplateGuid", permissionTemplateGuid);
    }

    public String getPermissionTemplateGuid()
    {
        return getString("PermissionTemplateGuid");
    }

    public boolean isReadReceiptRequested()
    {
        return getBoolean("ReadReceiptRequested");
    }

    public String getReceivedByEntryId()
    {
        return getString("ReceivedByEntryID");
    }

    public String getReceivedByName()
    {
        return getString("ReceivedByName");
    }

    public String getReceivedOnBehalfOfEntryId()
    {
        return getString("ReceivedOnBehalfOfEntryID");
    }

    public String getReceivedOnBehalfOfName()
    {
        return getString("ReceivedOnBehalfOfName");
    }

    /**
     * Returns the date and time at which the item was received.
     */
    public Date getReceivedTime()
    {
        return getDate("ReceivedTime");
    }

    public void setRecipientReassignmentProhibited(boolean recipientReassignmentProhibited)
    {
        put("RecipientReassignmentProhibited", recipientReassignmentProhibited);
    }

    public boolean isRecipientReassignmentProhibited()
    {
        return getBoolean("RecipientReassignmentProhibited");
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

    public void setRemoteStatus(OutlookRemoteStatus remoteStatus)
    {
        put("RemoteStatus", remoteStatus);
    }

    public OutlookRemoteStatus getRemoteStatus()
    {
        return getConstant("RemoteStatus", OutlookRemoteStatus.class);
    }

    public String getReplyRecipientNames()
    {
        return getString("ReplyRecipientNames");
    }

    /**
     * Returns all the reply recipient objects for the Outlook item.
     */
    public OutlookRecipients getReplyRecipients()
    {
        return new OutlookRecipients(getDispatch("ReplyRecipients"));
    }

    /**
     * Returns the date when the MailItem object expires, after which the Messaging Records Management (MRM) Assistant will delete the item.
     */
    public Date getRetentionExpirationDate()
    {
        return getDate("RetentionExpirationDate");
    }

    /**
     * Returns the name of the retention policy.
     */
    public String getRetentionPolicyName()
    {
        return getString("RetentionPolicyName");
    }

    /**
     * Sets the folder in which a copy of the email message will be saved after being sent.
     *
     * @param folder the folder in which a copy of the email message will be saved after being sent
     */
    public void setSaveSentMessageFolder(OutlookFolder folder)
    {
        put("SaveSentMessageFolder", folder);
    }

    /**
     * Returns the folder in which a copy of the email message will be saved after being sent.
     */
    public OutlookFolder getSaveSentMessageFolder()
    {
        return new OutlookFolder(getDispatch("SaveSentMessageFolder"));
    }

    public void setSender(OutlookAddressEntry sender)
    {
        put("Sender", sender);
    }

    public OutlookAddressEntry getSender()
    {
        return new OutlookAddressEntry(getDispatch("Sender"));
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
     * Indicates if a message has been sent.
     */
    public boolean isSent()
    {
        return getBoolean("Sent");
    }

    /**
     * Returns the date and time on which the Outlook item was sent.
     */
    public Date getSentOn()
    {
        return getDate("SentOn");
    }

    public void setSentOnBehalfOfName(String sentOnBehalfOfName)
    {
        put("SentOnBehalfOfName", sentOnBehalfOfName);
    }

    public String getSentOnBehalfOfName()
    {
        return getString("SentOnBehalfOfName");
    }

    /**
     * Determines if the item has been submitted.
     */
    public boolean isSubmitted()
    {
        return getBoolean("Submitted");
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

    public void setTo(String to)
    {
        put("To", to);
    }

    public String getTo()
    {
        return getString("To");
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

    public void setVotingOptions(String votingOptions)
    {
        put("VotingOptions", votingOptions);
    }

    public String getVotingOptions()
    {
        return getString("VotingOptions");
    }

    public void setVotingResponse(String votingResponse)
    {
        put("VotingResponse", votingResponse);
    }

    public String getVotingResponse()
    {
        return getString("VotingResponse");
    }
}
