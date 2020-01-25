package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookItemType;
import com.dtxmaker.jcom.outlook.constant.OutlookShowItemCount;
import com.jacob.com.Dispatch;

/**
 * Represents an Outlook folder.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.folder">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.folder</a>
 */
public class OutlookFolder extends Outlook
{
    OutlookFolder(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Adds a Microsoft Exchange public folder to the public folder's Favorites folder.
     */
    public void addToPublicFolderFavorites()
    {
        call("AddToPFFavorites");
    }

    /**
     * Copies the current folder in its entirety to the destination folder.
     *
     * @param destinationFolder Folder object that represents the destination folder.
     */
    public void copyTo(OutlookFolder destinationFolder)
    {
        call("CopyTo", destinationFolder);
    }

    /**
     * Deletes an object from the collection.
     */
    public void delete()
    {
        call("Delete");
    }

    /**
     * Displays a new Explorer object for the folder.
     */
    public void display()
    {
        call("Display");
    }

    /**
     * Returns a Folder object of the requested name.
     *
     * @param name The name of folder to return.
     * @return A Folder object that represents the folder of the requested name.
     */
    public OutlookFolder getFolder(String name)
    {
        return new OutlookFolder(callDispatch("Folders", name));
    }

    /**
     * Moves a folder to the specified destination folder.
     *
     * @param destinationFolder The destination Folder for the Folder that is being moved.
     */
    public void moveTo(OutlookFolder destinationFolder)
    {
        call("MoveTo", destinationFolder);
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    public void setAddressBookName(String addressBookName)
    {
        put("AddressBookName", addressBookName);
    }

    public String getAddressBookName()
    {
        return getString("AddressBookName");
    }

    public boolean isCustomViewsOnly()
    {
        return getBoolean("CustomViewsOnly");
    }

    public OutlookItemType getDefaultItemType()
    {
        return getConstant("DefaultItemType", OutlookItemType.class);
    }

    public String getDefaultMessageClass()
    {
        return getString("DefaultMessageClass");
    }

    public void setDescription(String description)
    {
        put("Description", description);
    }

    public String getDescription()
    {
        return getString("Description");
    }

    public String getEntryId()
    {
        return getString("EntryID");
    }

    public String getFolderPath()
    {
        return getString("FolderPath");
    }

    public OutlookFolders getFolders()
    {
        return new OutlookFolders(getDispatch("Folders"));
    }

    public void setInAppFolderSyncObject(boolean inAppFolderSyncObject)
    {
        put("InAppFolderSyncObject", inAppFolderSyncObject);
    }

    public boolean isInAppFolderSyncObject()
    {
        return getBoolean("InAppFolderSyncObject");
    }

    public boolean isSharePointFolder()
    {
        return getBoolean("IsSharePointFolder");
    }

    public OutlookItems getItems()
    {
        return new OutlookItems(getDispatch("Items"));
    }

    public String getName()
    {
        return getString("Name");
    }

    public void setName(String name)
    {
        put("Name", name);
    }

    public void setShowAsAddressBook(boolean showAsAddressBook)
    {
        put("ShowAsOutlookAB", showAsAddressBook);
    }

    public boolean isShowAsAddressBook()
    {
        return getBoolean("ShowAsOutlookAB");
    }

    public void setShowItemCount(OutlookShowItemCount showItemCount)
    {
        put("ShowItemCount", showItemCount);
    }

    public OutlookShowItemCount getShowItemCount()
    {
        return getConstant("ShowItemCount", OutlookShowItemCount.class);
    }

    public String getStoreId()
    {
        return getString("StoreID");
    }

    public int getUnReadItemCount()
    {
        return getInt("UnReadItemCount");
    }

    public void setWebViewOn(boolean webViewOn)
    {
        put("WebViewOn", webViewOn);
    }

    public boolean isWebViewOn()
    {
        return getBoolean("WebViewOn");
    }

    public void setWebViewUrl(String webViewUrl)
    {
        put("WebViewURL", webViewUrl);
    }

    public String getWebViewUrl()
    {
        return getString("WebViewURL");
    }
}
