package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.Countable;
import com.dtxmaker.jcom.outlook.constant.OutlookObjectClass;
import com.jacob.com.Dispatch;
import lombok.SneakyThrows;

import java.util.ArrayList;
import java.util.List;

public class OutlookFolder extends Outlook implements Countable<OutlookItem>
{
    public OutlookFolder(OutlookApplication application, Dispatch dispatch)
    {
        super(application, dispatch);
    }

    private Dispatch getItemsDispatch()
    {
        return cache.computeIfAbsent("Items", key -> Dispatch.get(dispatch, key).getDispatch());
    }

    private Dispatch getFoldersDispatch()
    {
        return cache.computeIfAbsent("Folders", key -> Dispatch.get(dispatch, key).getDispatch());
    }

    private int getCount(Dispatch objects)
    {
        return Dispatch.get(objects, "Count").getInt();
    }

    public String getName()
    {
        return getString("Name");
    }

    public void setName(String name)
    {
        put("Name", name);
    }

    public int getFolderCount()
    {
        return getCount(getFoldersDispatch());
    }

    /**
     * Return all sub folders in this folder.
     *
     * @return sub folders in this folder.
     */
    public OutlookFolder[] getFolders()
    {
        int count = getFolderCount();
        OutlookFolder[] folders = new OutlookFolder[count];

        for (int i = 0; i < count; i++)
        {
            folders[i] = getFolderAt(i + 1);
        }

        return folders;
    }

    public OutlookFolder getFolderAt(int index)
    {
        Dispatch folder = Dispatch.call(getFoldersDispatch(), "Item", index).toDispatch();
        return new OutlookFolder(application, folder);
    }

    public OutlookFolder getFolder(String name)
    {
        Dispatch folder = Dispatch.call(dispatch, "Folders", name).toDispatch();
        return new OutlookFolder(application, folder);
    }

    public OutlookFolder addFolder(String name)
    {
        Dispatch folder = Dispatch.call(getFoldersDispatch(), "Add", name).toDispatch();
        return new OutlookFolder(application, folder);
    }

    public void removeAllFolders()
    {
        for (int index = getFolderCount(); index > 0; index--)
        {
            removeFolderAt(index);
        }
    }

    public void removeFolderAt(int index)
    {
        Dispatch.call(getFoldersDispatch(), "Remove", index);
    }

    public void removeFolder(String name)
    {
        Dispatch.call(getFoldersDispatch(), "Remove", name);
    }

    @Override
    public int getItemCount()
    {
        return getCount(getItemsDispatch());
    }

    @Override
    public OutlookItem[] getItems()
    {
        int count = getItemCount();
        OutlookItem[] items = new OutlookItem[count];

        for (int i = 0; i < count; i++)
        {
            items[i] = getItemAt(i + 1);
        }

        return items;
    }

    public <T extends OutlookItem> List<T> getItems(Class<T> type)
    {
        int count = getItemCount();
        List<T> items = new ArrayList<>(count);
        OutlookObjectClass objectClass = OutlookObjectClass.findByType(type);

        for (int index = 1; index <= count; index++)
        {
            T item = getItemAt(index, type);

            if (objectClass == null || item.getObjectClass() == objectClass.getValue())
            {
                items.add(item);
            }
        }

        return items;
    }

    @Override
    public OutlookItem getItemAt(int index)
    {
        Dispatch item = Dispatch.call(getItemsDispatch(), "Item", index).toDispatch();
        return new OutlookItem(application, item);
    }

    @SneakyThrows
    public <T extends OutlookItem> T getItemAt(int index, Class<T> type)
    {
        Dispatch item = Dispatch.call(getItemsDispatch(), "Item", index).toDispatch();
        return type.getConstructor(OutlookApplication.class, Dispatch.class).newInstance(application, item);
    }

    @Override
    public void removeAllItems()
    {
        for (int index = getItemCount(); index > 0; index--)
        {
            removeItemAt(index);
        }
    }

    @Override
    public void removeItemAt(int index)
    {
        Dispatch.call(getItemsDispatch(), "Remove", index);
    }
}
