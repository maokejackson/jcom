package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.ImmutableList;
import com.jacob.com.Dispatch;

import java.util.Date;

/**
 * A collection of TimeZone objects.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.timezones">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.timezones</a>
 */
public class OutlookTimeZones extends Outlook implements ImmutableList<OutlookTimeZone>
{
    OutlookTimeZones(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Converts a date/time value from one time zone to another time zone.
     *
     * @param sourceDateTime      A date/time value expressed in the original time zone.
     * @param sourceTimeZone      The original time zone of the date/time value that is to be converted.
     * @param destinationTimeZone The target time zone to which the date/time value is to be converted.
     * @return A Date value that represents the date and time expressed in the DestinationTimeZone.
     */
    public Date convertTime(Date sourceDateTime, OutlookTimeZone sourceTimeZone, OutlookTimeZone destinationTimeZone)
    {
        return callDate("ConvertTime", sourceDateTime, sourceTimeZone, destinationTimeZone);
    }

    /**
     * Returns a TimeZone object from the collection.
     *
     * @param index An Integer representing a 1-based index into the TimeZones collection.
     * @return A TimeZone object that represents the specified object in the collection.
     */
    @Override
    public OutlookTimeZone getItem(int index)
    {
        return new OutlookTimeZone(callDispatch("Item", index));
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    @Override
    public int getCount()
    {
        return getInt("Count");
    }

    /**
     * @return a TimeZone value that represents the current Windows system local time zone.
     */
    public OutlookTimeZone getCurrentTimeZone()
    {
        return new OutlookTimeZone(getDispatch("CurrentTimeZone"));
    }
}
