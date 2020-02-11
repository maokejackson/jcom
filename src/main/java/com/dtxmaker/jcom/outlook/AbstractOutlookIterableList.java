package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.IterableList;
import com.jacob.com.Dispatch;

import java.util.ConcurrentModificationException;
import java.util.Iterator;
import java.util.NoSuchElementException;
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

    @Override
    public Iterator<T> iterator()
    {
        return new Itr(getCount());
    }

    private class Itr implements Iterator<T>
    {
        private final int size;
        private       int cursor;

        private Itr(int size)
        {
            this.size = size;
        }

        @Override
        public boolean hasNext()
        {
            return cursor != size;
        }

        @Override
        public T next()
        {
            int index = cursor + 1;
            if (index > size) throw new NoSuchElementException();
            if (index > getCount()) throw new ConcurrentModificationException();
            cursor = index;
            return getItem(index);
        }
    }
}
