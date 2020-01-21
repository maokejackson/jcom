package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookDefaultFolder;
import com.dtxmaker.jcom.outlook.constant.OutlookObjectClass;
import com.jacob.com.Dispatch;

import static com.dtxmaker.jcom.outlook.constant.OutlookDefaultFolder.DELETED_ITEMS;

public class OutlookItem extends Outlook
{
    public OutlookItem(OutlookApplication application, Dispatch dispatch)
    {
        super(application, dispatch);
    }

    public void setCategories(String categories)
    {
        put("Categories", categories);
    }

    public String getCategories()
    {
        return getString("Categories");
    }

    /**
     * Attach a file to this item.
     *
     * @param filePath the absolute path of the file.
     */
    public void attachFile(String filePath)
    {
        getAttachments().add(filePath);
    }

    /**
     * Get attachments of this item.
     *
     * @return attachments of this item.
     */
    public OutlookAttachments getAttachments()
    {
        Dispatch dispatch = Dispatch.get(this.dispatch, "Attachments").getDispatch();
        return new OutlookAttachments(application, dispatch);
    }

    /**
     * Save item in its default folder.
     */
    public void save()
    {
        Dispatch.call(dispatch, "Save");
    }

    /**
     * Move item to specific <code>folder</code>.
     *
     * @param folder the destination folder
     */
    public void moveTo(OutlookFolder folder)
    {
        Dispatch.call(dispatch, "Move", folder.getDispatch());
    }

    /**
     * Move item to specific <code>folder</code>.
     *
     * @param folder the destination folder
     */
    public void moveTo(OutlookDefaultFolder folder)
    {
        moveTo(application.getFolder(folder));
    }

    /**
     * Move item to Deleted Items folder.
     */
    public void delete()
    {
        Dispatch.call(dispatch, "Move", application.getFolder(DELETED_ITEMS));
    }

    /**
     * Test if this item belongs to specific object class.
     *
     * @param objectClass the desired object class to check
     * @return <code>true</code> when this item belongs to specific object class.
     */
    public boolean isObject(OutlookObjectClass objectClass)
    {
        return getObjectClass() == objectClass.getValue();
    }
}
