package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

/**
 * Contains a set of Attachment objects that represent the attachments in an Outlook item.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.attachments">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.attachments</a>
 */
public class OutlookAttachments extends AbstractOutlookMutableList<OutlookAttachment, String>
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
    @Override
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
    @Override
    public OutlookAttachment getItem(int index)
    {
        return new OutlookAttachment(callDispatch("Item", index));
    }
}
