package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

/**
 * Contains a collection of Recipient objects for an Outlook item.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.recipients">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.recipients</a>
 */
public class OutlookRecipients extends AbstractOutlookMutableList<OutlookRecipient, String>
{
    OutlookRecipients(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Creates a new recipient in the Recipients collection.
     *
     * @param name The name of the recipient; it can be a string representing the display name, the alias, or the full SMTP email address of the recipient.
     * @return A Recipient object that represents the new recipient.
     */
    @Override
    public OutlookRecipient add(String name)
    {
        return new OutlookRecipient(callDispatch("Add", name));
    }

    /**
     * Returns a Recipient object from the collection.
     *
     * @param index Either the index number of the object, or a value used to match the default property of an object in the collection.
     * @return A Recipient object that represents the specified object.
     */
    @Override
    public OutlookRecipient getItem(int index)
    {
        return new OutlookRecipient(callDispatch("Item", index));
    }

    /**
     * Attempts to resolve all the Recipient objects in the Recipients collection against the Address Book.
     *
     * @return <cdde>true</cdde> if all of the objects were resolved, <code>false</code> if one or more were not.
     */
    public boolean resolveAll()
    {
        return callBoolean("ResolveAll");
    }
}
