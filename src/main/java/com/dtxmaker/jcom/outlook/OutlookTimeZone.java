package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

import java.util.Date;

/**
 * Represents information for a time zone as supported by Windows.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.timezone">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.timezone</a>
 */
public class OutlookTimeZone extends Outlook
{
    OutlookTimeZone(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    /**
     * Returns a int value that represents the difference in minutes of between the local time in this time zone and the Coordinated Universal Time (UTC).
     */
    public int getBias()
    {
        return getInt("Bias");
    }

    /**
     * Returns a int value that represents the time offset in minutes from the Bias to account for daylight time in this time zone.
     */
    public int getDaylightBias()
    {
        return getInt("DaylightBias");
    }

    /**
     * Returns a Date value that represents the date and time in this time zone when time changes over to daylight time in the current year.
     */
    public Date getDaylightDate()
    {
        return getDate("DaylightDate");
    }

    /**
     * Returns a String that identifies the time zone in daylight time.
     */
    public String getDaylightDesignation()
    {
        return getString("DaylightDesignation");
    }

    /**
     * Returns a String that uniquely identifies the time zone.
     */
    public String getId()
    {
        return getString("ID");
    }

    /**
     * Returns a String that represents the identifier of the time zone.
     */
    public String getName()
    {
        return getString("Name");
    }

    /**
     * Returns a int value that represents the time offset in minutes from the Bias to account for standard time in this time zone.
     */
    public int getStandardBias()
    {
        return getInt("StandardBias");
    }

    /**
     * Returns a Date value that represents the date and time in this time zone when time changes over to standard time.
     */
    public Date getStandardDate()
    {
        return getDate("StandardDate");
    }

    /**
     * Returns a String that identifies the time zone in standard time.
     */
    public String getStandardDesignation()
    {
        return getString("StandardDesignation");
    }
}
