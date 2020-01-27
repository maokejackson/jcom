package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

/**
 * Represents a journal entry in a Journal folder.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.journalitem">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.journalitem</a>
 */
public class OutlookJournal extends OutlookItem
{
    OutlookJournal(Dispatch dispatch)
    {
        super(dispatch);
    }
}
