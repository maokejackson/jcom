package com.dtxmaker.jcom;

import com.dtxmaker.jcom.util.EnumUtils;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.util.Date;
import java.util.Objects;

public abstract class Base
{
    private final Dispatch dispatch;

    protected Base(Dispatch dispatch)
    {
        this.dispatch = Objects.requireNonNull(dispatch);
    }

    protected final Dispatch getDispatch()
    {
        return dispatch;
    }

    private Object[] convertObjects(Object... objects)
    {
        if (objects == null) return new Object[0];

        Object[] out = new Object[objects.length];

        for (int i = 0; i < objects.length; i++)
        {
            Object obj = objects[i];

            if (obj instanceof Constant)
            {
                out[i] = ((Constant) obj).getValue();
            }
            else if (obj instanceof Base)
            {
                out[i] = ((Base) obj).dispatch;
            }
            else
            {
                out[i] = obj;
            }
        }

        return out;
    }

    /* *****************************************************
     *                                                     *
     *                     Method Call                     *
     *                                                     *
     *******************************************************/

    protected final Variant call(String method, Object... args)
    {
        return Dispatch.call(dispatch, method, convertObjects(args));
    }

    protected final Dispatch callDispatch(String method, Object... args)
    {
        return call(method, args).getDispatch();
    }

    protected final String callString(String method, Object... args)
    {
        return call(method, args).getString();
    }

    protected final int callInt(String method, Object... args)
    {
        return call(method, args).getInt();
    }

    protected final long callLong(String method, Object... args)
    {
        return call(method, args).getLong();
    }

    protected final Date callDate(String method, Object... args)
    {
        return call(method, args).getJavaDate();
    }

    protected final boolean callBoolean(String method, Object... args)
    {
        return call(method, args).getBoolean();
    }

    /* *****************************************************
     *                                                     *
     *                     Set Property                    *
     *                                                     *
     *******************************************************/

    protected final void put(String name, String value)
    {
        Dispatch.put(dispatch, name, value == null ? "" : value);
    }

    protected final void put(String name, Constant value)
    {
        Dispatch.put(dispatch, name, value.getValue());
    }

    protected final void put(String name, Base value)
    {
        Dispatch.put(dispatch, name, value.getDispatch());
    }

    protected final void put(String name, Object value)
    {
        Dispatch.put(dispatch, name, value);
    }

    /* *****************************************************
     *                                                     *
     *                     Get Property                    *
     *                                                     *
     *******************************************************/

    protected final Variant get(String name)
    {
        return Dispatch.get(dispatch, name);
    }

    protected final Dispatch getDispatch(String name)
    {
        return get(name).getDispatch();
    }

    protected final String getString(String name)
    {
        return get(name).getString();
    }

    protected final int getInt(String name)
    {
        return get(name).getInt();
    }

    protected final long getLong(String name)
    {
        return get(name).getLong();
    }

    protected final Date getDate(String name)
    {
        return get(name).getJavaDate();
    }

    protected final boolean getBoolean(String name)
    {
        return get(name).getBoolean();
    }

    protected final <T extends Enum<T> & Constant> T getConstant(String name, Class<T> type)
    {
        return EnumUtils.findByValue(type, getInt(name));
    }
}
