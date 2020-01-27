package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

/**
 * Represents a change to the recipient's Calendar folder initiated by another party or as a result of a group action.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.meetingitem">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.meetingitem</a>
 */
public class OutlookMeeting extends OutlookItem
{
    OutlookMeeting(Dispatch dispatch)
    {
        super(dispatch);
    }
}
