package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookPermission;
import com.dtxmaker.jcom.outlook.constant.OutlookPermissionService;
import com.jacob.com.Dispatch;

public class OutlookItem extends Outlook
{
    OutlookItem(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Move item to specific <code>folder</code>.
     *
     * @param folder the destination folder
     */
    public void moveTo(OutlookFolder folder)
    {
        call("Move", folder.getDispatch());
    }

    /**
     * Save item in its default folder.
     */
    public void save()
    {
        call("Save");
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    public void setCategories(String categories)
    {
        put("Categories", categories);
    }

    public String getCategories()
    {
        return getString("Categories");
    }

    /**
     * Get attachments of this item.
     *
     * @return attachments of this item.
     */
    public OutlookAttachments getAttachments()
    {
        return new OutlookAttachments(getDispatch("Attachments"));
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
