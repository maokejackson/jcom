package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

/**
 * Represents the collection of Category objects that define the Master Category List for a namespace.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.categories">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.categories</a>
 */
public class OutlookCategories extends AbstractOutlookMutableList<OutlookCategory>
{
    OutlookCategories(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Creates a new Category object and appends it to the Categories collection.
     *
     * @param name The name of the new category.
     * @return A Category object that represents the new category.
     */
    public OutlookCategory add(String name)
    {
        return new OutlookCategory(callDispatch("Add", name));
    }

    /**
     * Returns a Category object from the collection.
     *
     * @param index A Long value representing the index number of the object in the collection.
     * @return A Category object that represents the specified object.
     */
    @Override
    public OutlookCategory getItem(int index)
    {
        return new OutlookCategory(callDispatch("Item", index));
    }

    /**
     * Returns a Category object from the collection.
     *
     * @param name A String value representing either the Name or CategoryID property value of an object in the collection.
     * @return A Category object that represents the specified object.
     */
    public OutlookCategory getItem(String name)
    {
        return new OutlookCategory(callDispatch("Item", name));
    }

    /**
     * Removes an object from the collection.
     *
     * @param name A String value representing either the Name or CategoryID property value of an object in the collection.
     */
    public void remove(String name)
    {
        call("Remove", name);
    }
}
