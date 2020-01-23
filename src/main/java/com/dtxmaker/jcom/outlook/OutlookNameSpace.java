package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookDefaultFolderType;
import com.jacob.com.Dispatch;

public class OutlookNameSpace extends Outlook
{
    OutlookNameSpace(OutlookApplication application)
    {
        super(application, Dispatch.call(application.getDispatch(), "GetNamespace", "MAPI").toDispatch());
    }

    private Dispatch getDefaultFolder(int defaultFolder)
    {
        return Dispatch.call(dispatch, "GetDefaultFolder", defaultFolder).toDispatch();
    }

    public OutlookDefaultFolder getDefaultFolder(OutlookDefaultFolderType defaultFolder)
    {
        return new OutlookDefaultFolder(application, getDefaultFolder(defaultFolder.getValue()));
    }

    public String getType()
    {
        return getString("Type");
    }

    public void sendAndReceive(boolean showProgressDialog)
    {
        Dispatch.call(dispatch, "SendAndReceive", showProgressDialog);
    }
}
