package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookObjectClass;
import com.jacob.com.Dispatch;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

abstract class Outlook
{
    final OutlookApplication application;
    final Dispatch           dispatch;

    final Map<String, Dispatch> cache = new HashMap<>();

    Outlook(Dispatch dispatch)
    {
        this(null, dispatch);
    }

    Outlook(OutlookApplication application, Dispatch dispatch)
    {
        this.application = application;
        this.dispatch = Objects.requireNonNull(dispatch);
    }

    public final OutlookObjectClass getObjectClass()
    {
        int value = Dispatch.get(dispatch, "Class").getInt();
        return OutlookObjectClass.findByValue(value);
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

    final boolean getBoolean(String name)
    {
        return Dispatch.get(dispatch, name).getBoolean();
    }
}
