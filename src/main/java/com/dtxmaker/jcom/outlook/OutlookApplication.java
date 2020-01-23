package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.library.LanguageSettings;
import com.dtxmaker.jcom.outlook.constant.OutlookDefaultFolderType;
import com.dtxmaker.jcom.outlook.constant.OutlookItemType;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComFailException;
import com.jacob.com.Dispatch;

import static com.dtxmaker.jcom.outlook.constant.OutlookItemType.CONTACT;
import static com.dtxmaker.jcom.outlook.constant.OutlookItemType.MAIL;

/**
 * https://docs.microsoft.com/en-us/office/vba/api/overview/outlook
 */
public class OutlookApplication extends Outlook
{
    private static final String OUTLOOK_APPLICATION = "Outlook.Application";

    private final Dispatch namespace;

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
        super(new ActiveXComponent(OUTLOOK_APPLICATION));
        namespace = Dispatch.call(dispatch, "GetNamespace", "MAPI").toDispatch();
    }

    public String getDefaultProfileName()
    {
        return getString("DefaultProfileName");
    }

    public boolean isTrusted()
    {
        return getBoolean("IsTrusted");
    }

    public LanguageSettings getLanguageSettings()
    {
        return new LanguageSettings(Dispatch.call(dispatch, "LanguageSettings").toDispatch());
    }

    public String getName()
    {
        return getString("Name");
    }

    public String getProductCode()
    {
        return getString("ProductCode");
    }

    public String getVersion()
    {
        return getString("Version");
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
