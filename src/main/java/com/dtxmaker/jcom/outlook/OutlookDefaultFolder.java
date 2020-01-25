package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

public class OutlookDefaultFolder extends OutlookFolder
{
    OutlookDefaultFolder(Dispatch dispatch)
    {
        super(dispatch);
    }

    @Override
    public final void setName(String name)
    {
        // Default folder cannot rename
    }
}
