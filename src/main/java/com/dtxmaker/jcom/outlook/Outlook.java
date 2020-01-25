package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.Base;
import com.dtxmaker.jcom.outlook.constant.OutlookObjectClass;
import com.jacob.com.Dispatch;

abstract class Outlook extends Base
{
    Outlook(Dispatch dispatch)
    {
        super(dispatch);
    }

    public final OutlookObjectClass getObjectClass()
    {
        return getConstant("Class", OutlookObjectClass.class);
    }

    public final OutlookNameSpace getSession()
    {
        return new OutlookNameSpace(callDispatch("GetNamespace", "MAPI"));
    }
}
