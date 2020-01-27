package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

import java.util.Date;

/**
 * Represents a change to the recipient's Calendar folder initiated by another party or as a result of a group action.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.meetingitem">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.meetingitem</a>
 */
public class OutlookMeeting extends AbstractOutlookInternalItem
{
    OutlookMeeting(Dispatch dispatch)
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
     * Indicates whether the MeetingItem represents the latest version of the item on the organizer's calendar.
     */
    public boolean isLatestVersion()
    {
        return getBoolean("IsLatestVersion");
    }

    /**
     * Returns the URL for the Meeting Workspace that the meeting item is linked to.
     */
    public String getMeetingWorkspaceURL()
    {
        return getString("MeetingWorkspaceURL");
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

    /**
     * Returns the date and time at which the item was received.
     */
    public Date getReceivedTime()
    {
        return getDate("ReceivedTime");
    }

    /**
     * Returns all the recipients for the Outlook item.
     */
    public final OutlookRecipients getRecipients()
    {
        return new OutlookRecipients(callDispatch("Recipients"));
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

    public void setReminderTime(Date reminderTime)
    {
        put("ReminderTime", reminderTime);
    }

    public Date getReminderTime()
    {
        return getDate("ReminderTime");
    }

    /**
     * Returns all the reply recipient objects for the Outlook item.
     */
    public OutlookRecipients getReplyRecipients()
    {
        return new OutlookRecipients(getDispatch("ReplyRecipients"));
    }

    /**
     * Returns the date when the MeetingItem object expires, after which the Messaging Records Management (MRM) Assistant will delete the item.
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

    /**
     * Determines if the item has been submitted.
     */
    public boolean isSubmitted()
    {
        return getBoolean("Submitted");
    }
}
