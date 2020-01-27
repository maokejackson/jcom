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
public class OutlookMail extends OutlookItem
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

    public void setAutoForwarded(boolean autoForwarded)
    {
        put("AutoForwarded", autoForwarded);
    }

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

    public void setBillingInformation(String billingInformation)
    {
        put("BillingInformation", billingInformation);
    }

    public String getBillingInformation()
    {
        return getString("BillingInformation");
    }

    public void setBodyFormat(OutlookBodyFormat bodyFormat)
    {
        put("BodyFormat", bodyFormat);
    }

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

    public void setCompanies(String companies)
    {
        put("Companies", companies);
    }

    public String getCompanies()
    {
        return getString("Companies");
    }

    public void setDeferredDeliveryTime(Date deferredDeliveryTime)
    {
        put("DeferredDeliveryTime", deferredDeliveryTime);
    }

    public Date getDeferredDeliveryTime()
    {
        return getDate("DeferredDeliveryTime");
    }

    public void setDeleteAfterSubmit(boolean deleteAfterSubmit)
    {
        put("DeleteAfterSubmit", deleteAfterSubmit);
    }

    public boolean isDeleteAfterSubmit()
    {
        return getBoolean("DeleteAfterSubmit");
    }

    public OutlookDownloadState getDownloadState()
    {
        return getConstant("DownloadState", OutlookDownloadState.class);
    }

    public void setExpiryTime(Date expiryTime)
    {
        put("ExpiryTime", expiryTime);
    }

    public Date getExpiryTime()
    {
        return getDate("ExpiryTime");
    }

    public void setFlagRequest(String flagRequest)
    {
        put("FlagRequest", flagRequest);
    }

    public String getFlagRequest()
    {
        return getString("FlagRequest");
    }

    public void setHtmlBody(String htmlBody)
    {
        put("HTMLBody", htmlBody);
    }

    public String getHtmlBody()
    {
        return getString("HTMLBody");
    }

    public void setInternetCodePage(long internetCodePage)
    {
        put("InternetCodepage", internetCodePage);
    }

    public long getInternetCodePage()
    {
        return getLong("InternetCodepage");
    }

    public boolean isMarkedAsTask()
    {
        return getBoolean("IsMarkedAsTask");
    }

    public void setOriginatorDeliveryReportRequested(boolean originatorDeliveryReportRequested)
    {
        put("OriginatorDeliveryReportRequested", originatorDeliveryReportRequested);
    }

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

    public Date getRetentionExpirationDate()
    {
        return getDate("RetentionExpirationDate");
    }

    public String getRetentionPolicyName()
    {
        return getString("RetentionPolicyName");
    }

    public void setSaveSentMessageFolder(OutlookFolder folder)
    {
        put("SaveSentMessageFolder", folder);
    }

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

    public String getSenderEmailAddress()
    {
        return getString("SenderEmailAddress");
    }

    public String getSenderEmailType()
    {
        return getString("SenderEmailType");
    }

    public String getSenderName()
    {
        return getString("SenderName");
    }

    public boolean isSent()
    {
        return getBoolean("Sent");
    }

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

    public boolean isSubmitted()
    {
        return getBoolean("Submitted");
    }

    public void setTaskCompletedDate(Date taskCompletedDate)
    {
        put("TaskCompletedDate", taskCompletedDate);
    }

    public Date getTaskCompletedDate()
    {
        return getDate("TaskCompletedDate");
    }

    public void setTaskDueDate(Date taskDueDate)
    {
        put("TaskDueDate", taskDueDate);
    }

    public Date getTaskDueDate()
    {
        return getDate("TaskDueDate");
    }

    public void setTaskStartDate(Date taskStartDate)
    {
        put("TaskStartDate", taskStartDate);
    }

    public Date getTaskStartDate()
    {
        return getDate("TaskStartDate");
    }

    public void setTaskSubject(String taskSubject)
    {
        put("TaskSubject", taskSubject);
    }

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

    public void setToDoTaskOrdinal(Date toDoTaskOrdinal)
    {
        put("ToDoTaskOrdinal", toDoTaskOrdinal);
    }

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
