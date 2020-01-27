package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.IterableList;
import com.jacob.com.Dispatch;

import java.util.Optional;

public abstract class AbstractOutlookIterableList<T extends Outlook, V> extends AbstractOutlookMutableList<T, V>
        implements IterableList<T>
{
    AbstractOutlookIterableList(Dispatch dispatch)
    {
        super(dispatch);
    }

    abstract T newInstance(Dispatch dispatch);

    T getItem(String method, Object... args)
    {
        return Optional.ofNullable(call(method, args))
                .filter(variant -> !variant.isNull())
                .map(variant -> newInstance(variant.getDispatch()))
                .orElse(null);
    }

    @Override
    public final T getFirst()
    {
        return getItem("GetFirst");
    }

    @Override
    public final T getLast()
    {
        return getItem("GetLast");
    }

    @Override
    public final T getNext()
    {
        return getItem("GetNext");
    }

    @Override
    public final T getPrevious()
    {
        return getItem("GetPrevious");
    }
}
