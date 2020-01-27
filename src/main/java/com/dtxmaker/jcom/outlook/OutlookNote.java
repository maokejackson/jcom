package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

/**
 * Represents a note in a Notes folder.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.noteitem">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.noteitem</a>
 */
public class OutlookNote extends OutlookItem
{
    OutlookNote(Dispatch dispatch)
    {
        super(dispatch);
    }
}
