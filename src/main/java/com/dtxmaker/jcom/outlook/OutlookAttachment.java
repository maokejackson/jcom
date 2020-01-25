package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookAttachmentBlockLevel;
import com.dtxmaker.jcom.outlook.constant.OutlookAttachmentType;
import com.jacob.com.Dispatch;

/**
 * Represents a document or link to a document contained in an Outlook item.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.attachment">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.attachment</a>
 */
public class OutlookAttachment extends Outlook
{
    OutlookAttachment(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Deletes an object from the collection.
     */
    public void delete()
    {
        call("Delete");
    }

    /**
     * Returns the full path to the attached file that is in a temporary files folder.
     */
    public String getTemporaryFilePath()
    {
        return callString("GetTemporaryFilePath");
    }

    /**
     * Saves the attachment to the specified path.
     *
     * @param path The location at which to save the attachment.
     */
    public void saveAsFile(String path)
    {
        call("SaveAsFile", path);
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    /**
     * Returns an enumeration that specifies if there is any restriction on the attachment based on its file extension.
     */
    public OutlookAttachmentBlockLevel getBlockLevel()
    {
        return getConstant("BlockLevel", OutlookAttachmentBlockLevel.class);
    }

    /**
     * Sets the name, which does not need to be the actual file name, displayed below the icon representing the embedded attachment.
     *
     * @param displayName The display name.
     */
    public void setDisplayName(String displayName)
    {
        put("DisplayName", displayName);
    }

    /**
     * Returns the name, which does not need to be the actual file name, displayed below the icon representing the embedded attachment.
     */
    public String getDisplayName()
    {
        return getString("DisplayName");
    }

    /**
     * Returns the file name of the attachment.
     */
    public String getFileName()
    {
        return getString("FileName");
    }

    /**
     * Returns the position of the object within the collection.
     */
    public int getIndex()
    {
        return getInt("Index");
    }

    /**
     * Returns the full path to the linked attached file.
     */
    public String getPathName()
    {
        return getString("PathName");
    }

    /**
     * Sets the position of the attachment within the body of the item.
     *
     * @param position The position of the attachment.
     */
    public void setPosition(int position)
    {
        put("Position", position);
    }

    /**
     * Returns the position of the attachment within the body of the item.
     */
    public int getPosition()
    {
        return getInt("Position");
    }

    /**
     * Returns the size (in bytes) of the attachment.
     */
    public long getSize()
    {
        return getLong("Size");
    }

    /**
     * Returns an enumeration indicating the type of the specified object.
     */
    public OutlookAttachmentType getType()
    {
        return getConstant("Type", OutlookAttachmentType.class);
    }
}
