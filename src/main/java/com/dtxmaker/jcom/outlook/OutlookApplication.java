package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookDefaultFolderType;
import com.dtxmaker.jcom.outlook.constant.OutlookItemType;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComFailException;
import com.jacob.com.Dispatch;

import java.util.HashMap;
import java.util.Map;

import static com.dtxmaker.jcom.outlook.constant.OutlookItemType.CONTACT;
import static com.dtxmaker.jcom.outlook.constant.OutlookItemType.MAIL;

public class OutlookApplication
{
    private static final String OUTLOOK_APPLICATION = "Outlook.Application";

    private final Dispatch dispatch;
    private final Dispatch namespace;

    private final Map<String, Dispatch> cache = new HashMap<>();

    public static boolean isInstalled()
    {
        try
        {
            new ActiveXComponent(OUTLOOK_APPLICATION);
            return true;
        }
        catch (ComFailException e)
        {
            return false;
        }
    }

    public OutlookApplication()
    {
        ActiveXComponent component = new ActiveXComponent(OUTLOOK_APPLICATION);
        dispatch = component.getObject();
        namespace = Dispatch.call(dispatch, "GetNamespace", "MAPI").toDispatch();
    }

    private Dispatch createItem(OutlookItemType itemType)
    {
        return Dispatch.call(dispatch, "CreateItem", itemType.getValue()).toDispatch();
    }

    private Dispatch getDefaultFolder(int defaultFolder)
    {
        return cache.computeIfAbsent("GetDefaultFolder" + defaultFolder,
                key -> Dispatch.call(namespace, "GetDefaultFolder", defaultFolder).toDispatch());
    }

    public OutlookMail createMail()
    {
        return new OutlookMail(this, createItem(MAIL));
    }

    public OutlookContact createContact()
    {
        return new OutlookContact(this, createItem(CONTACT));
    }

    public OutlookDefaultFolder getDefaultFolder(OutlookDefaultFolderType defaultFolder)
    {
        return new OutlookDefaultFolder(this, getDefaultFolder(defaultFolder.getValue()));
    }
}
