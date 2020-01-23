package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

public class OutlookDefaultFolder extends OutlookFolder
{
    public OutlookDefaultFolder(OutlookApplication application, Dispatch dispatch)
    {
        super(application, dispatch);
    }

    @Override
    public final void setName(String name)
    {
        // Default folder cannot rename
    }
}
