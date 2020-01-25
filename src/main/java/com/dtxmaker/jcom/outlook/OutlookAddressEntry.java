package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookAddressEntryUserType;
import com.dtxmaker.jcom.outlook.constant.OutlookDisplayType;
import com.jacob.com.Dispatch;

import java.util.Date;
import java.util.Optional;

/**
 * Represents a person, group, or public folder to which the messaging system can deliver messages.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.addressentry.type">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.addressentry.type</a>
 */
public class OutlookAddressEntry extends Outlook
{
    OutlookAddressEntry(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Deletes an object from the collection.
     */
    public void delete()
    {
        call("Delete");
    }

    /**
     * Returns a ContactItem object that represents the AddressEntry, if the AddressEntry corresponds to a contact in an Outlook Contacts Address Book (CAB).
     *
     * @return A ContactItem object that corresponds to the AddressEntry. Returns <code>null</code> if the AddressEntry object does not correspond to a contact in a Contacts Address Book.
     */
    public OutlookContact getContact()
    {
        return Optional.ofNullable(call("GetContact"))
                .filter(variant -> !variant.isNull())
                .map(variant -> new OutlookContact(variant.getDispatch()))
                .orElse(null);
    }

    /**
     * Returns a String value that represents the availability of the individual user for a period of 30 days from the start date, beginning at midnight of the date specified.
     *
     * @param start      the date.
     * @param minPerChar the length of each time slot in minutes.
     * @return A String value that represents the availability of the user for the specified period. The string value contains one character for each time slot within the specified period.
     */
    public String getFreeBusy(Date start, int minPerChar)
    {
        return getFreeBusy(start, minPerChar, false);
    }

    /**
     * Returns a String value that represents the availability of the individual user for a period of 30 days from the start date, beginning at midnight of the date specified.
     *
     * @param start          the date.
     * @param minPerChar     the length of each time slot in minutes.
     * @param completeFormat a Boolean value that represents the level of information returned for each time slot.
     * @return A String value that represents the availability of the user for the specified period. The string value contains one character for each time slot within the specified period.
     */
    public String getFreeBusy(Date start, int minPerChar, boolean completeFormat)
    {
        return callString("GetFreeBusy", start, minPerChar, completeFormat);
    }

    /**
     * Posts a change to the AddressEntry object in the messaging system.
     */
    public void update()
    {
        update(true, false);
    }

    /**
     * Posts a change to the AddressEntry object in the messaging system.
     *
     * @param makePermanent A value of True indicates that the property cache is flushed and all changes are committed in the underlying address book. A value of False indicates that the property cache is flushed but not committed to persistent storage.
     * @param refresh       A value of True indicates that the property cache is reloaded from the values in the underlying address book. A value of False indicates that the property cache is not reloaded.
     */
    public void update(boolean makePermanent, boolean refresh)
    {
        call("Update", makePermanent, refresh);
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    /**
     * Sets a String representing the email address of the AddressEntry.
     *
     * @param address the email address of the AddressEntry.
     */
    public void setAddress(String address)
    {
        put("Address", address);
    }

    /**
     * Returns a String representing the email address of the AddressEntry.
     */
    public String getAddress()
    {
        return getString("Address");
    }

    /**
     * Returns an enumeration representing the user type of the {@link OutlookAddressEntry}.
     */
    public OutlookAddressEntryUserType getAddressEntryUserType()
    {
        return getConstant("AddressEntryUserType", OutlookAddressEntryUserType.class);
    }

    /**
     * Returns an enumeration that describes the nature of the {@link OutlookAddressEntry}.
     */
    public OutlookDisplayType getDisplayType()
    {
        return getConstant("DisplayType", OutlookDisplayType.class);
    }

    /**
     * Returns a String representing the unique identifier for the object.
     */
    public String getId()
    {
        return getString("ID");
    }

    /**
     * Sets a String value that represents the display name for the object.
     *
     * @param name the display name for the object.
     */
    public void setName(String name)
    {
        put("Name", name);
    }

    /**
     * Returns a String value that represents the display name for the object.
     */
    public String getName()
    {
        return getString("Name");
    }

    /**
     * Sets a String representing the type of entry for this address such as an Internet Address, MacMail Address, or Microsoft Mail Address.
     *
     * @param type the type of entry for this address
     */
    public void setType(String type)
    {
        put("Type", type);
    }

    /**
     * Returns a String representing the type of entry for this address such as an Internet Address, MacMail Address, or Microsoft Mail Address.
     */
    public String getType()
    {
        return getString("Type");
    }
}
