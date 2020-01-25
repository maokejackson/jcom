package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

/**
 * Contains a set of Attachment objects that represent the attachments in an Outlook item.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.attachments">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.attachments</a>
 */
public class OutlookAttachments extends Outlook
{
    OutlookAttachments(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Creates a new attachment in the Attachments collection.
     *
     * @param source The source of the attachment. This can be a file (represented by the full file system path with a file name) or an Outlook item that constitutes the attachment.
     * @return An Attachment object that represents the new attachment.
     */
    public OutlookAttachment add(String source)
    {
        return new OutlookAttachment(callDispatch("Add", source));
    }

    /**
     * Returns an object from the collection.
     *
     * @param index Either the index number of the object, or a value used to match the default property of an object in the collection.
     * @return An Attachment object that represents the specified object.
     */
    public OutlookAttachment getItem(int index)
    {
        return new OutlookAttachment(callDispatch("Item", index));
    }

    /**
     * Removes an object from the collection.
     *
     * @param index The 1-based index value of the object within the collection.
     */
    public void remove(int index)
    {
        call("Remove", index);
    }

    /**
     * Remove all objects from the collection.
     */
    public void removeAll()
    {
        for (int index = getCount(); index >= 1; index--)
        {
            remove(index);
        }
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    /**
     * Returns the count of objects in the specified collection.
     */
    public int getCount()
    {
        return getInt("Count");
    }
}
