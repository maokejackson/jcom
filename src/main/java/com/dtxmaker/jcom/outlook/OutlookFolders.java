package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookDefaultFolderType;
import com.jacob.com.Dispatch;

/**
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.folders
 */
public class OutlookFolders extends AbstractOutlookIterableList<OutlookFolder>
{
    OutlookFolders(Dispatch dispatch)
    {
        super(dispatch);
    }

    @Override
    OutlookFolder newInstance(Dispatch dispatch)
    {
        return new OutlookFolder(dispatch);
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
     * Returns a Folder object from the collection.
     *
     * @param index Either the index number of the object, or a value used to match the default property of an object in the collection.
     * @return A Folder object that represents the specified object.
     */
    @Override
    public OutlookFolder getItem(int index)
    {
        return new OutlookFolder(callDispatch("Item", index));
    }
}
