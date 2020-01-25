package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookDefaultFolderType;
import com.jacob.com.Dispatch;

import java.util.Optional;

/**
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.folders
 */
public class OutlookFolders extends Outlook
{
    OutlookFolders(Dispatch dispatch)
    {
        super(dispatch);
    }

    private OutlookFolder getFolder(String method)
    {
        return Optional.ofNullable(call(method))
                .filter(variant -> !variant.isNull())
                .map(variant -> new OutlookFolder(variant.getDispatch()))
                .orElse(null);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Creates a new folder in the Folders collection. The new folder will default to the same type as the folder in which it is created.
     *
     * @param name The display name for the new folder.
     * @return A Folder object that represents the new folder.
     */
    public OutlookFolder add(String name)
    {
        return new OutlookFolder(callDispatch("Add", name));
    }

    /**
     * Creates a new folder in the Folders collection.
     *
     * @param name The display name for the new folder.
     * @param type The Outlook folder type for the new folder.
     * @return A Folder object that represents the new folder.
     */
    public OutlookFolder add(String name, OutlookDefaultFolderType type)
    {
        return new OutlookFolder(callDispatch("Add", name, type));
    }

    /**
     * Returns the first object in the Folders collection.
     *
     * @return A Folder object that represents the first object contained by the collection.
     */
    public OutlookFolder getFirst()
    {
        return getFolder("GetFirst");
    }

    /**
     * Returns the last object in the Folders collection.
     *
     * @return A Folder object that represents the last object contained by the collection.
     */
    public OutlookFolder getLast()
    {
        return getFolder("GetLast");
    }

    /**
     * Returns the next object in the Folders collection.
     *
     * @return A Folder object that represents the next object contained by the collection.
     */
    public OutlookFolder getNext()
    {
        return getFolder("GetNext");
    }

    /**
     * Returns the previous object in the Folders collection.
     *
     * @return A Folder object that represents the previous object contained by the collection.
     */
    public OutlookFolder getPrevious()
    {
        return getFolder("GetPrevious");
    }

    /**
     * Returns a Folder object from the collection.
     *
     * @param index Either the index number of the object, or a value used to match the default property of an object in the collection.
     * @return A Folder object that represents the specified object.
     */
    public OutlookFolder getItem(int index)
    {
        return new OutlookFolder(callDispatch("Item", index));
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
