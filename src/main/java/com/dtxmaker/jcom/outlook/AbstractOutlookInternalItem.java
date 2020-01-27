package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookImportance;
import com.dtxmaker.jcom.outlook.constant.OutlookSensitivity;
import com.jacob.com.Dispatch;

public abstract class AbstractOutlookInternalItem extends OutlookItem
{
    AbstractOutlookInternalItem(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Displays the Show Categories dialog box, which allows you to select categories that correspond to the subject of the item.
     */
    public final void showCategoriesDialog()
    {
        call("ShowCategoriesDialog");
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    /**
     * Returns all the attachments for the specified item.
     */
    public final OutlookAttachments getAttachments()
    {
        return new OutlookAttachments(getDispatch("Attachments"));
    }

    /**
     * Sets the billing information associated with the Outlook item.
     *
     * @param billingInformation the billing information
     */
    public final void setBillingInformation(String billingInformation)
    {
        put("BillingInformation", billingInformation);
    }

    /**
     * Returns the billing information associated with the Outlook item.
     */
    public final String getBillingInformation()
    {
        return getString("BillingInformation");
    }

    /**
     * Sets the names of the companies associated with the Outlook item.
     *
     * @param companies the names of the companies
     */
    public final void setCompanies(String companies)
    {
        put("Companies", companies);
    }

    /**
     * Returns the names of the companies associated with the Outlook item.
     */
    public final String getCompanies()
    {
        return getString("Companies");
    }

    /**
     * Returns a String that uniquely identifies a Conversation object that the Outlook item belongs to.
     */
    public final String getConversationId()
    {
        return getString("ConversationID");
    }

    /**
     * Returns the relative position of the item within the conversation thread.
     */
    public final String getConversationIndex()
    {
        return getString("ConversationIndex");
    }

    /**
     * Returns the topic of the conversation thread of the Outlook item.
     */
    public final String getConversationTopic()
    {
        return getString("ConversationTopic");
    }

    /**
     * Sets the relative importance level for the Outlook item.
     *
     * @param importance The relative importance level for the Outlook item.
     */
    public final void setImportance(OutlookImportance importance)
    {
        put("Importance", importance);
    }

    /**
     * Returns the relative importance level for the Outlook item.
     */
    public final OutlookImportance getImportance()
    {
        return getConstant("Importance", OutlookImportance.class);
    }

    /**
     * Sets the mileage for an item.
     *
     * @param mileage The mileage.
     */
    public final void setMileage(String mileage)
    {
        put("Mileage", mileage);
    }

    /**
     * Returns the mileage for an item.
     */
    public final String getMileage()
    {
        return getString("Mileage");
    }

    /**
     * Sets <code>true</code> to not age the Outlook item.
     *
     * @param noAging <code>true</code> to not age the Outlook item; Otherwise, <code>false</code>.
     */
    public final void setNoAging(boolean noAging)
    {
        put("NoAging", noAging);
    }

    /**
     * Returns <code>true</code> to not age the Outlook item.
     */
    public final boolean isNoAging()
    {
        return getBoolean("NoAging");
    }

    /**
     * Returns the build number of the Outlook application for an Outlook item.
     */
    public final long getOutlookInternalVersion()
    {
        return getLong("OutlookInternalVersion");
    }

    /**
     * Returns the major and minor version number of the Outlook application for an Outlook item.
     */
    public final String getOutlookVersion()
    {
        return getString("OutlookVersion");
    }

    /**
     * Sets the sensitivity for the Outlook item.
     *
     * @param sensitivity the sensitivity
     */
    public final void setSensitivity(OutlookSensitivity sensitivity)
    {
        put("Sensitivity", sensitivity);
    }

    /**
     * Returns the sensitivity for the Outlook item.
     */
    public final OutlookSensitivity getSensitivity()
    {
        return getConstant("Sensitivity", OutlookSensitivity.class);
    }

    /**
     * Sets if the Outlook item has not been opened (read).
     *
     * @param unRead <code>true</code> if the Outlook item has not been opened (read)
     */
    public final void setUnRead(boolean unRead)
    {
        put("UnRead", unRead);
    }

    /**
     * Returns <code>true</code> if the Outlook item has not been opened (read).
     */
    public final boolean isUnRead()
    {
        return getBoolean("UnRead");
    }
}
