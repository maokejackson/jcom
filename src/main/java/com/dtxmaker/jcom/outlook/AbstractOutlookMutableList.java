package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.MutableList;
import com.jacob.com.Dispatch;

public abstract class AbstractOutlookMutableList<T extends Outlook> extends Outlook implements MutableList<T>
{
    AbstractOutlookMutableList(Dispatch dispatch)
    {
        super(dispatch);
    }

    @Override
    public final void remove(int index)
    {
        call("Remove", index);
    }

    @Override
    public final void removeAll()
    {
        for (int index = getCount(); index >= 1; index--)
        {
            remove(index);
        }
    }

    @Override
    public final int getCount()
    {
        return getInt("Count");
    }
}
