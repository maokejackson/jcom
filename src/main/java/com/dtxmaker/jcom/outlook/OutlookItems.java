package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookItemType;
import com.jacob.com.Dispatch;
import lombok.SneakyThrows;

import java.util.Optional;

/**
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.items
 */
public class OutlookItems extends Outlook
{
    OutlookItems(Dispatch dispatch)
    {
        super(dispatch);
    }

    private OutlookItem getItem(String method, Object... args)
    {
        return Optional.ofNullable(call(method, args))
                .filter(variant -> !variant.isNull())
                .map(variant -> new OutlookItem(variant.getDispatch()))
                .orElse(null);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Creates a new Outlook item in the Items collection for the folder.
     *
     * @return An Object value that represents the new Outlook item.
     */
    public OutlookItem add()
    {
        return new OutlookItem(callDispatch("Add"));
    }

    /**
     * Creates a new Outlook item in the Items collection for the folder.
     *
     * @param type The Outlook item type for the new item.
     * @param <T>  The Outlook item type.
     * @return An Object value that represents the new Outlook item.
     */
    @SneakyThrows
    public <T extends OutlookItem> T add(Class<T> type)
    {
        int value = OutlookItemType.findValueByType(type);
        Dispatch item = callDispatch("Add", value);
        return type.getConstructor(Dispatch.class).newInstance(item);
    }

    /**
     * Locates and returns a Microsoft Outlook item object that satisfies the given Filter.
     *
     * @param filter A string that specifies the criteria that the returned object must satisfy.
     * @return An Object value that represents an Outlook item if the call succeeds; returns <code>null</code> if it fails.
     */
    public OutlookItem find(String filter)
    {
        return getItem("Find", filter);
    }

    /**
     * After the Find method runs, this method finds and returns the next Outlook item in the specified collection.
     *
     * @return An Object value that represents the next Outlook item found in the collection.
     */
    public OutlookItem findNext()
    {
        return getItem("FindNext");
    }

    /**
     * Returns the first object in the collection.
     *
     * @return An Object value that represents the first object contained by the collection.
     */
    public OutlookItem getFirst()
    {
        return getItem("GetFirst");
    }

    /**
     * Returns the last object in the collection.
     *
     * @return An Object value that represents the last object contained by the collection.
     */
    public OutlookItem getLast()
    {
        return getItem("GetLast");
    }

    /**
     * Returns the next object in the collection.
     *
     * @return An Object value that represents the next object contained by the collection.
     */
    public OutlookItem getNext()
    {
        return getItem("GetNext");
    }

    /**
     * Returns the previous object in the collection.
     *
     * @return An Object value that represents the previous object contained by the collection.
     */
    public OutlookItem getPrevious()
    {
        return getItem("GetPrevious");
    }

    /**
     * Returns an Outlook item from a collection.
     *
     * @param index Either the index number of the object, or a value used to match the default property of an object in the collection.
     * @return An Object value that represents the specified object.
     */
    public OutlookItem getItem(int index)
    {
        return new OutlookItem(callDispatch("Item", index));
    }

    /**
     * Removes an object from the collection.
     *
     * @param index The 1-based index value of the object within the collection.
     */
    public void remove(int index)
    {
        call("Remove", index);
    }

    /**
     * Remove all objects from the collection.
     */
    public void removeAll()
    {
        for (int index = getCount(); index >= 1; index--)
        {
            remove(index);
        }
    }

    /**
     * Clears the properties that have been cached with the SetColumns method.
     */
    public void resetColumns()
    {
        call("ResetColumns");
    }

    /**
     * Applies a filter to the Items collection, returning a new collection containing all of the items from the original that match the filter.
     *
     * @param filter A filter string expression to be applied. For details, see the Find method.
     * @return An Items collection that represents the items from the original Items collection which match the filter.
     */
    public OutlookItems restrict(String filter)
    {
        return new OutlookItems(callDispatch("Restrict", filter));
    }

    /**
     * Caches certain properties for extremely fast access to those particular properties of each item in an Items collection.
     *
     * @param columns A string that contains the names of the properties to cache. The property names are delimited by commas in this string.
     */
    public void setColumns(String columns)
    {
        call("SetColumns", columns);
    }

    /**
     * Sorts the collection of items descending by the specified property. The index for the collection is reset to 1 upon completion of this method.
     *
     * @param property The name of the property by which to sort, which may be enclosed in brackets, for example, "[CompanyName]".
     */
    public void sort(String property)
    {
        call("Sort", property);
    }

    /**
     * Sorts the collection of items by the specified property. The index for the collection is reset to 1 upon completion of this method.
     *
     * @param property   The name of the property by which to sort, which may be enclosed in brackets, for example, "[CompanyName]".
     * @param descending <code>true</code> to sort in descending order.
     */
    public void sort(String property, boolean descending)
    {
        call("Sort", property, descending);
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    public int getCount()
    {
        return getInt("Count");
    }
}
