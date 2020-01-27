package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

/**
 * Represents a post in a public folder that others may browse.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.postitem">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.postitem</a>
 */
public class OutlookPost extends OutlookItem
{
    OutlookPost(Dispatch dispatch)
    {
        super(dispatch);
    }
}
