package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.outlook.constant.OutlookObjectClass;
import com.dtxmaker.jcom.util.EnumUtils;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.util.Date;
import java.util.Objects;

abstract class Outlook
{
    private final Dispatch dispatch;

    Outlook(Dispatch dispatch)
    {
        this.dispatch = Objects.requireNonNull(dispatch);
    }

    public final OutlookObjectClass getObjectClass()
    {
        return getConstant("Class", OutlookObjectClass.class);
    }

    public final OutlookNameSpace getSession()
    {
        return new OutlookNameSpace(Dispatch.call(dispatch, "GetNamespace", "MAPI").getDispatch());
    }

    final Dispatch getDispatch()
    {
        return dispatch;
    }

    final Variant call(String method, Object... args)
    {
        return Dispatch.call(dispatch, method, args);
    }

    final Dispatch callDispatch(String method, Object... args)
    {
        return call(method, args).getDispatch();
    }

    final String callString(String method, Object... args)
    {
        return call(method, args).getString();
    }

    final int callInt(String method, Object... args)
    {
        return call(method, args).getInt();
    }

    final long callLong(String method, Object... args)
    {
        return call(method, args).getLong();
    }

    final Date callDate(String method, Object... args)
    {
        return call(method, args).getJavaDate();
    }

    final boolean callBoolean(String method, Object... args)
    {
        return call(method, args).getBoolean();
    }

    final void put(String name, String value)
    {
        Dispatch.put(dispatch, name, value == null ? "" : value);
    }

    final void put(String name, Constant value)
    {
        Dispatch.put(dispatch, name, value.getValue());
    }

    final void put(String name, Outlook value)
    {
        put(name, value.getDispatch());
    }

    final void put(String name, Object value)
    {
        Dispatch.put(dispatch, name, value);
    }

    final Dispatch getDispatch(String name)
    {
        return Dispatch.get(dispatch, name).getDispatch();
    }

    final String getString(String name)
    {
        return Dispatch.get(dispatch, name).getString();
    }

    final int getInt(String name)
    {
        return Dispatch.get(dispatch, name).getInt();
    }

    final long getLong(String name)
    {
        return Dispatch.get(dispatch, name).getLong();
    }

    final Date getDate(String name)
    {
        return Dispatch.get(dispatch, name).getJavaDate();
    }

    final boolean getBoolean(String name)
    {
        return Dispatch.get(dispatch, name).getBoolean();
    }

    final <T extends Enum<T> & Constant> T getConstant(String name, Class<T> type)
    {
        return EnumUtils.findByValue(type, getInt(name));
    }
}
