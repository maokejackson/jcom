package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookDefaultFolderType;
import com.dtxmaker.jcom.outlook.constant.OutlookPermission;
import com.dtxmaker.jcom.outlook.constant.OutlookPermissionService;
import com.jacob.com.Dispatch;

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
        return new OutlookAttachments(application, getDispatch("Attachments"));
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
     * Move item to designated <code>default folder</code>.
     *
     * @param defaultFolder the designated default folder
     */
    public void moveTo(OutlookDefaultFolderType defaultFolder)
    {
        moveTo(application.getDefaultFolder(defaultFolder));
    }

    /**
     * Move item to Deleted Items folder.
     */
    public void delete()
    {
        moveTo(OutlookDefaultFolderType.DELETED_ITEMS);
    }

    public int getOutlookInternalVersion()
    {
        return getInt("OutlookInternalVersion");
    }

    public String getOutlookVersion()
    {
        return getString("OutlookVersion");
    }

    public void setPermission(OutlookPermission permission)
    {
        put("Permission", permission);
    }

    public OutlookPermission getPermission()
    {
        return getConstant("Permission", OutlookPermission.class);
    }

    public void setPermissionService(OutlookPermissionService permissionService)
    {
        put("PermissionService", permissionService);
    }

    public OutlookPermissionService getPermissionService()
    {
        return getConstant("PermissionService", OutlookPermissionService.class);
    }

    public void setPermissionTemplateGuid(String permissionTemplateGuid)
    {
        put("PermissionTemplateGuid", permissionTemplateGuid);
    }

    public String getPermissionTemplateGuid()
    {
        return getString("PermissionTemplateGuid");
    }
}
