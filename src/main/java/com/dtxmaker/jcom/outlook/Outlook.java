package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

abstract class Outlook
{
    final OutlookApplication application;
    final Dispatch           dispatch;

    final Map<String, Dispatch> cache = new HashMap<>();

    Outlook(OutlookApplication application, Dispatch dispatch)
    {
        this.application = application;
        this.dispatch = dispatch;
    }

    final OutlookApplication getApplication()
    {
        return application;
    }

    final Dispatch getDispatch()
    {
        return dispatch;
    }

    final void put(String name, String value)
    {
        if (value == null) value = "";
        Dispatch.put(dispatch, name, value);
    }

    final void put(String name, int value)
    {
        Dispatch.put(dispatch, name, value);
    }

    final void put(String name, Date value)
    {
        Dispatch.put(dispatch, name, value);
    }

    final int getObjectClass()
    {
        return Dispatch.get(dispatch, "Class").getInt();
    }

    final String getString(String name)
    {
        return Dispatch.get(dispatch, name).getString();
    }

    final int getInt(String name)
    {
        return Dispatch.get(dispatch, name).getInt();
    }

    final Date getDate(String name)
    {
        return Dispatch.get(dispatch, name).getJavaDate();
    }
}
