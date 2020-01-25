package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

/**
 * Represents a user-defined category by which Outlook items can be grouped.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.category">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.category</a>
 */
public class OutlookCategory extends Outlook
{
    OutlookCategory(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    /**
     * Returns a String value that represents the unique identifier for the Category object.
     */
    public String getId()
    {
        return getString("CategoryID");
    }

    /**
     * Sets a String value that represents the display name for the object.
     *
     * @param name The display name for the object.
     */
    public void setName(String name)
    {
        put("Name", name);
    }

    /**
     * Returns a String value that represents the display name for the object.
     */
    public String getName()
    {
        return getString("Name");
    }
}
