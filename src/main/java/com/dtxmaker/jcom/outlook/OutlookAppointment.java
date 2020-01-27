package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

/**
 * Represents a meeting, a one-time appointment, or a recurring appointment or meeting in the Calendar folder.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.appointmentitem">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.appointmentitem</a>
 */
public class OutlookAppointment extends OutlookItem
{
    OutlookAppointment(Dispatch dispatch)
    {
        super(dispatch);
    }
}
