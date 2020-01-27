package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

/**
 * Represents a task (an assigned, delegated, or self-imposed task to be performed within a specified time frame) in a Tasks folder.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.taskitem">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.taskitem</a>
 */
public class OutlookTask extends OutlookItem
{
    OutlookTask(Dispatch dispatch)
    {
        super(dispatch);
    }
}
