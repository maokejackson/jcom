package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.Countable;
import com.jacob.com.Dispatch;

public class OutlookAttachments extends Outlook implements Countable<Dispatch>
{
    public OutlookAttachments(OutlookApplication application, Dispatch dispatch)
    {
        super(application, dispatch);
    }

    public void add(String filePath)
    {
        Dispatch.call(dispatch, "Add", filePath);
    }

    @Override
    public int getItemCount()
    {
        return getInt("Count");
    }

    @Override
    public Dispatch[] getItems()
    {
        int count = getItemCount();
        Dispatch[] array = new Dispatch[count];
        for (int i = 0; i < count; i++)
        {
            array[i] = getItemAt(i + 1);
        }
        return array;
    }

    @Override
    public Dispatch getItemAt(int index)
    {
        return Dispatch.call(dispatch, "Item", index).toDispatch();
    }

    @Override
    public void removeAllItems()
    {
        int count = getItemCount();

        for (int index = 1; index <= count; index++)
        {
            removeItemAt(index);
        }
    }

    @Override
    public void removeItemAt(int index)
    {
        Dispatch.call(dispatch, "Remove", index);
    }
}
