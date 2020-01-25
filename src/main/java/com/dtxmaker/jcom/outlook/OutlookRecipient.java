package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookDisplayType;
import com.dtxmaker.jcom.outlook.constant.OutlookResponseStatus;
import com.dtxmaker.jcom.outlook.constant.OutlookTrackingStatus;
import com.jacob.com.Dispatch;

import java.util.Date;

/**
 * Represents a user or resource in Outlook, generally a mail or mobile message addressee.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.recipient">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.recipient</a>
 */
public class OutlookRecipient extends Outlook
{
    OutlookRecipient(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Deletes an object from the collection.
     */
    public void delete()
    {
        call("Delete");
    }

    /**
     * Returns free/busy information for the recipient.
     *
     * @param start      The start date for the returned period of free/busy information.
     * @param minPerChar The number of minutes per character represented in the returned free/busy string.
     * @return A String value that represents the free/busy information.
     */
    public String getFreeBusy(Date start, int minPerChar)
    {
        return getFreeBusy(start, minPerChar, false);
    }

    /**
     * Returns free/busy information for the recipient.
     *
     * @param start          The start date for the returned period of free/busy information.
     * @param minPerChar     The number of minutes per character represented in the returned free/busy string.
     * @param completeFormat <code>true</code> if the returned string should contain not only free/busy information, but also values for each character according to the {@link com.dtxmaker.jcom.outlook.constant.OutlookBusyStatus OutlookBusyStatus} constants.
     * @return A String value that represents the free/busy information.
     */
    public String getFreeBusy(Date start, int minPerChar, boolean completeFormat)
    {
        return callString("GetFreeBusy", start, minPerChar, completeFormat);
    }

    /**
     * Attempts to resolve a Recipient object against the Address Book.
     *
     * @return <code>true</code> if the object was resolved; otherwise, <code>false</code>.
     */
    public boolean resolve()
    {
        return callBoolean("Resolve");
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    public String getAddress()
    {
        return getString("Address");
    }

    public void setAddressEntry(OutlookAddressEntry addressEntry)
    {
        put("AddressEntry", addressEntry.getDispatch());
    }

    public OutlookAddressEntry getAddressEntry()
    {
        return new OutlookAddressEntry(getDispatch("AddressEntry"));
    }

    public void setAutoResponse(String autoResponse)
    {
        put("AutoResponse", autoResponse);
    }

    public String getAutoResponse()
    {
        return getString("AutoResponse");
    }

    public OutlookDisplayType getDisplayType()
    {
        return getConstant("DisplayType", OutlookDisplayType.class);
    }

    public String getEntryId()
    {
        return getString("EntryID");
    }

    public int getIndex()
    {
        return getInt("Index");
    }

    public OutlookResponseStatus getMeetingResponseStatus()
    {
        return getConstant("MeetingResponseStatus", OutlookResponseStatus.class);
    }

    public String getName()
    {
        return getString("Name");
    }

    public boolean isResolved()
    {
        return getBoolean("Resolved");
    }

    public void setSendable(boolean sendable)
    {
        put("Sendable", sendable);
    }

    public boolean isSendable()
    {
        return getBoolean("Sendable");
    }

    public void setTrackingStatus(OutlookTrackingStatus trackingStatus)
    {
        put("TrackingStatus", trackingStatus);
    }

    public OutlookTrackingStatus getTrackingStatus()
    {
        return getConstant("TrackingStatus", OutlookTrackingStatus.class);
    }

    public void setTrackingStatusTime(Date trackingStatusTime)
    {
        put("TrackingStatusTime", trackingStatusTime);
    }

    public Date getTrackingStatusTime()
    {
        return getDate("TrackingStatusTime");
    }

    public void setType(int type)
    {
        put("Type", type);
    }

    public int getType()
    {
        return getInt("Type");
    }
}
