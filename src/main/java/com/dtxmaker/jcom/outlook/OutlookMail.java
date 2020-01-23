package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.*;
import com.jacob.com.Dispatch;

import java.util.Date;

public class OutlookMail extends OutlookItem
{
    public OutlookMail(OutlookApplication application, Dispatch dispatch)
    {
        super(application, dispatch);
    }

    public void setSubject(String subject)
    {
        put("Subject", subject);
    }

    public String getSubject()
    {
        return getString("Subject");
    }

    public String getSenderName()
    {
        return getString("SenderName");
    }

    public void setSenderEmailAddress(String senderEmailAddress)
    {
        put("SenderEmailAddress", senderEmailAddress);
    }

    public String getSenderEmailAddress()
    {
        return getString("SenderEmailAddress");
    }

    public String getSenderEmailType()
    {
        return getString("SenderEmailType");
    }

    public void setTo(String to)
    {
        put("To", to);
    }

    public String getTo()
    {
        return getString("To");
    }

    public void setCc(String cc)
    {
        put("CC", cc);
    }

    public String getCc()
    {
        return getString("CC");
    }

    public void setBcc(String bcc)
    {
        put("BCC", bcc);
    }

    public String getBcc()
    {
        return getString("BCC");
    }

    public void setBody(String body)
    {
        put("Body", body);
    }

    public String getBody()
    {
        return getString("Body");
    }

    public void setHtmlBody(String htmlBody)
    {
        put("HTMLBody", htmlBody);
    }

    public String getHtmlBody()
    {
        return getString("HTMLBody");
    }

    public void setBodyFormat(OutlookBodyFormat bodyFormat)
    {
        put("BodyFormat", bodyFormat);
    }

    public OutlookBodyFormat getBodyFormat()
    {
        return getConstant("BodyFormat", OutlookBodyFormat.class);
    }

    public void setImportance(OutlookImportance importance)
    {
        put("Importance", importance);
    }

    public OutlookImportance getImportance()
    {
        return getConstant("Importance", OutlookImportance.class);
    }

    public void setAutoForwarded(boolean autoForwarded)
    {
        put("AutoForwarded", autoForwarded);
    }

    public boolean isAutoForwarded()
    {
        return getBoolean("AutoForwarded");
    }

    public boolean isAutoResolvedWinner()
    {
        return getBoolean("AutoResolvedWinner");
    }

    public void setBillingInformation(String billingInformation)
    {
        put("BillingInformation", billingInformation);
    }

    public String getBillingInformation()
    {
        return getString("BillingInformation");
    }

    public void setCompanies(String companies)
    {
        put("Companies", companies);
    }

    public String getCompanies()
    {
        return getString("Companies");
    }

    public String getConversationID()
    {
        return getString("ConversationID");
    }

    public String getConversationIndex()
    {
        return getString("ConversationIndex");
    }

    public String getConversationTopic()
    {
        return getString("ConversationTopic");
    }

    public Date getCreationTime()
    {
        return getDate("CreationTime");
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

    public String getEntryID()
    {
        return getString("EntryID");
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

    public void setInternetCodePage(int internetCodePage)
    {
        put("InternetCodepage", internetCodePage);
    }

    public int getInternetCodePage()
    {
        return getInt("InternetCodepage");
    }

    public boolean isConflict()
    {
        return getBoolean("IsConflict");
    }

    public boolean isMarkedAsTask()
    {
        return getBoolean("IsMarkedAsTask");
    }

    public Date getLastModificationTime()
    {
        return getDate("LastModificationTime");
    }

    public void setMarkForDownload(OutlookRemoteStatus remoteStatus)
    {
        put("MarkForDownload", remoteStatus);
    }

    public OutlookRemoteStatus getMarkForDownload()
    {
        return getConstant("MarkForDownload", OutlookRemoteStatus.class);
    }

    public void setMessageClass(String messageClass)
    {
        put("MessageClass", messageClass);
    }

    public String getMessageClass()
    {
        return getString("MessageClass");
    }

    public void setMileage(String mileage)
    {
        put("Mileage", mileage);
    }

    public String getMileage()
    {
        return getString("Mileage");
    }

    public void setNoAging(boolean noAging)
    {
        put("NoAging", noAging);
    }

    public boolean isNoAging()
    {
        return getBoolean("NoAging");
    }

    public void setOriginatorDeliveryReportRequested(boolean originatorDeliveryReportRequested)
    {
        put("OriginatorDeliveryReportRequested", originatorDeliveryReportRequested);
    }

    public boolean isOriginatorDeliveryReportRequested()
    {
        return getBoolean("OriginatorDeliveryReportRequested");
    }

    public boolean isReadReceiptRequested()
    {
        return getBoolean("ReadReceiptRequested");
    }

    public String getReceivedByEntryID()
    {
        return getString("ReceivedByEntryID");
    }

    public String getReceivedByName()
    {
        return getString("ReceivedByName");
    }

    public String getReceivedOnBehalfOfEntryID()
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

    public void setReminderOverrideDefault(String reminderOverrideDefault)
    {
        put("ReminderOverrideDefault", reminderOverrideDefault);
    }

    public String getReminderOverrideDefault()
    {
        return getString("ReminderOverrideDefault");
    }

    public void setReminderPlaySound(boolean reminderPlaySound)
    {
        put("ReminderPlaySound", reminderPlaySound);
    }

    public boolean isReminderPlaySound()
    {
        return getBoolean("ReminderPlaySound");
    }

    public void setReminderSet(boolean reminderSet)
    {
        put("ReminderSet", reminderSet);
    }

    public boolean getReminderSet()
    {
        return getBoolean("ReminderSet");
    }

    public void setReminderSoundFile(String reminderSoundFile)
    {
        put("ReminderSoundFile", reminderSoundFile);
    }

    public String getReminderSoundFile()
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

    public Date getRetentionExpirationDate()
    {
        return getDate("RetentionExpirationDate");
    }

    public String getRetentionPolicyName()
    {
        return getString("RetentionPolicyName");
    }

    public boolean isSaved()
    {
        return getBoolean("Saved");
    }

    public void setSaveSentMessageFolder(OutlookFolder folder)
    {
        put("SaveSentMessageFolder", folder.getDispatch());
    }

    public OutlookFolder getSaveSentMessageFolder()
    {
        return new OutlookFolder(application, getDispatch("SaveSentMessageFolder"));
    }

    public void setSensitivity(OutlookSensitivity sensitivity)
    {
        put("Sensitivity", sensitivity);
    }

    public OutlookSensitivity getSensitivity()
    {
        return getConstant("Sensitivity", OutlookSensitivity.class);
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

    public int getSize()
    {
        return getInt("Size");
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

    public void setToDoTaskOrdinal(Date toDoTaskOrdinal)
    {
        put("ToDoTaskOrdinal", toDoTaskOrdinal);
    }

    public Date getToDoTaskOrdinal()
    {
        return getDate("ToDoTaskOrdinal");
    }

    public void setUnRead(boolean unRead)
    {
        put("UnRead", unRead);
    }

    public boolean isUnRead()
    {
        return getBoolean("UnRead");
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

    /**
     * Send mail automatically.
     */
    public void send()
    {
        Dispatch.call(dispatch, "Send");
    }
}
