package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.*;
import com.jacob.com.Dispatch;

import java.util.Date;

/**
 * Represents a contact in a Contacts folder.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.contactitem">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.contactitem</a>
 */
public class OutlookContact extends OutlookItem
{
    OutlookContact(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Adds a logo picture to the current Electronic Business Card of the contact item.
     *
     * @param path The full path name that specifies the picture file to load.
     */
    public void addBusinessCardLogoPicture(String path)
    {
        call("AddBusinessCardLogoPicture", path);
    }

    /**
     * Adds a picture to a contact item.
     *
     * @param path The complete path and file name of the picture to be added to the contact item.
     */
    public void addPicture(String path)
    {
        call("AddPicture", path);
    }

    /**
     * Clears the ContactItem object as a task.
     */
    public void clearTaskFlag()
    {
        call("ClearTaskFlag");
    }

    /**
     * Creates a new MailItem object containing contact information and, optionally, an Electronic Business Card (EBC) image based on the specified ContactItem object.
     *
     * @return A MailItem object that represents the new email item containing the business card information.
     */
    public OutlookMail forwardAsBusinessCard()
    {
        return new OutlookMail(callDispatch("ForwardAsBusinessCard"));
    }

    /**
     * Creates a MailItem and attaches the contact information in vCard format.
     *
     * @return A MailItem object that represents the new mail item to which the contact information is attached.
     */
    public OutlookMail forwardAsVcard()
    {
        return new OutlookMail(callDispatch("ForwardAsVcard"));
    }

    /**
     * Marks it as a task and assigns a task interval for the object.
     *
     * @param markInterval the task interval
     */
    public void markAsTask(OutlookMarkInterval markInterval)
    {
        call("MarkAsTask", markInterval);
    }

    /**
     * Removes a picture from a Contact item.
     */
    public void removePicture()
    {
        call("RemovePicture");
    }

    /**
     * Resets the Electronic Business Card on the contact item to the default business card, deleting any custom layout and logo on the Electronic Business Card.
     */
    public void resetBusinessCard()
    {
        call("ResetBusinessCard");
    }

    /**
     * Saves an image of the business card generated from the specified ContactItem object.
     *
     * @param path The fully qualified path and file name of the image to be saved.
     */
    public void saveBusinessCardImage(String path)
    {
        call("SaveBusinessCardImage", path);
    }

    /**
     * Displays the electronic business card (EBC) editor dialog box for the ContactItem object.
     */
    public void showBusinessCardEditor()
    {
        call("ShowBusinessCardEditor");
    }

    /**
     * Displays the Check Address dialog box to verify address details of the contact.
     *
     * @param mailingAddress The type of address to be checked.
     */
    public void showCheckAddressDialog(OutlookMailingAddress mailingAddress)
    {
        call("ShowCheckAddressDialog", mailingAddress);
    }

    /**
     * Displays the Check Full Name dialog box to verify name details of the contact.
     */
    public void showCheckFullNameDialog()
    {
        call("ShowCheckFullNameDialog");
    }

    /**
     * Displays the Check Phone Number dialog box for a specified telephone number contained by a ContactItem object.
     *
     * @param phoneNumber The type of telephone number to be checked.
     */
    public void showCheckPhoneDialog(OutlookContactPhoneNumber phoneNumber)
    {
        call("ShowCheckPhoneDialog", phoneNumber);
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    public void setAccount(String account)
    {
        put("Account", account);
    }

    public String getAccount()
    {
        return getString("Account");
    }

    public void setAnniversary(Date anniversary)
    {
        put("Anniversary", anniversary);
    }

    public Date getAnniversary()
    {
        return getDate("Anniversary");
    }

    public void setAssistantName(String assistantName)
    {
        put("AssistantName", assistantName);
    }

    public String getAssistantName()
    {
        return getString("AssistantName");
    }

    public void setAssistantTelephoneNumber(String assistantTelephoneNumber)
    {
        put("AssistantTelephoneNumber", assistantTelephoneNumber);
    }

    public String getAssistantTelephoneNumber()
    {
        return getString("AssistantTelephoneNumber");
    }

    public void setBillingInformation(String billingInformation)
    {
        put("BillingInformation", billingInformation);
    }

    public String getBillingInformation()
    {
        return getString("BillingInformation");
    }

    public void setBirthday(Date birthday)
    {
        put("Birthday", birthday);
    }

    public Date getBirthday()
    {
        return getDate("Birthday");
    }

    public void setBusiness2TelephoneNumber(String business2TelephoneNumber)
    {
        put("Business2TelephoneNumber", business2TelephoneNumber);
    }

    public String getBusiness2TelephoneNumber()
    {
        return getString("Business2TelephoneNumber");
    }

    public void setBusinessAddress(String businessAddress)
    {
        put("BusinessAddress", businessAddress);
    }

    public String getBusinessAddress()
    {
        return getString("BusinessAddress");
    }

    public void setBusinessAddressCity(String businessAddressCity)
    {
        put("BusinessAddressCity", businessAddressCity);
    }

    public String getBusinessAddressCity()
    {
        return getString("BusinessAddressCity");
    }

    public void setBusinessAddressCountry(String businessAddressCountry)
    {
        put("BusinessAddressCountry", businessAddressCountry);
    }

    public String getBusinessAddressCountry()
    {
        return getString("BusinessAddressCountry");
    }

    public void setBusinessAddressPostalCode(String businessAddressPostalCode)
    {
        put("BusinessAddressPostalCode", businessAddressPostalCode);
    }

    public String getBusinessAddressPostalCode()
    {
        return getString("BusinessAddressPostalCode");
    }

    public void setBusinessAddressPostOfficeBox(String businessAddressPostOfficeBox)
    {
        put("BusinessAddressPostOfficeBox", businessAddressPostOfficeBox);
    }

    public String getBusinessAddressPostOfficeBox()
    {
        return getString("BusinessAddressPostOfficeBox");
    }

    public void setBusinessAddressState(String businessAddressState)
    {
        put("BusinessAddressState", businessAddressState);
    }

    public String getBusinessAddressState()
    {
        return getString("BusinessAddressState");
    }

    public void setBusinessAddressStreet(String businessAddressStreet)
    {
        put("BusinessAddressStreet", businessAddressStreet);
    }

    public String getBusinessAddressStreet()
    {
        return getString("BusinessAddressStreet");
    }

    public void setBusinessCardLayoutXml(String businessCardLayoutXml)
    {
        put("BusinessCardLayoutXml", businessCardLayoutXml);
    }

    public String getBusinessCardLayoutXml()
    {
        return getString("BusinessCardLayoutXml");
    }

    public OutlookBusinessCardType getBusinessCardType()
    {
        return getConstant("BusinessCardType", OutlookBusinessCardType.class);
    }

    public void setBusinessFaxNumber(String businessFaxNumber)
    {
        put("BusinessFaxNumber", businessFaxNumber);
    }

    public String getBusinessFaxNumber()
    {
        return getString("BusinessFaxNumber");
    }

    public void setBusinessHomePage(String businessHomePage)
    {
        put("BusinessHomePage", businessHomePage);
    }

    public String getBusinessHomePage()
    {
        return getString("BusinessHomePage");
    }

    public void setBusinessTelephoneNumber(String businessTelephoneNumber)
    {
        put("BusinessTelephoneNumber", businessTelephoneNumber);
    }

    public String getBusinessTelephoneNumber()
    {
        return getString("BusinessTelephoneNumber");
    }

    public void setCallbackTelephoneNumber(String callbackTelephoneNumber)
    {
        put("CallbackTelephoneNumber", callbackTelephoneNumber);
    }

    public String getCallbackTelephoneNumber()
    {
        return getString("CallbackTelephoneNumber");
    }

    public void setCarTelephoneNumber(String carTelephoneNumber)
    {
        put("CarTelephoneNumber", carTelephoneNumber);
    }

    public String getCarTelephoneNumber()
    {
        return getString("CarTelephoneNumber");
    }

    public void setChildren(String children)
    {
        put("Children", children);
    }

    public String getChildren()
    {
        return getString("Children");
    }

    public void setCompanies(String companies)
    {
        put("Companies", companies);
    }

    public String getCompanies()
    {
        return getString("Companies");
    }

    public String getCompanyAndFullName()
    {
        return getString("CompanyAndFullName");
    }

    public String getCompanyLastFirstNoSpace()
    {
        return getString("CompanyLastFirstNoSpace");
    }

    public String getCompanyLastFirstSpaceOnly()
    {
        return getString("CompanyLastFirstSpaceOnly");
    }

    public void setCompanyMainTelephoneNumber(String companyMainTelephoneNumber)
    {
        put("CompanyMainTelephoneNumber", companyMainTelephoneNumber);
    }

    public String getCompanyMainTelephoneNumber()
    {
        return getString("CompanyMainTelephoneNumber");
    }

    public void setCompanyName(String companyName)
    {
        put("CompanyName", companyName);
    }

    public String getCompanyName()
    {
        return getString("CompanyName");
    }

    public void setComputerNetworkName(String computerNetworkName)
    {
        put("ComputerNetworkName", computerNetworkName);
    }

    public String getComputerNetworkName()
    {
        return getString("ComputerNetworkName");
    }

    public void setCustomerId(String customerId)
    {
        put("CustomerID", customerId);
    }

    public String getCustomerId()
    {
        return getString("CustomerID");
    }

    public void setDepartment(String department)
    {
        put("Department", department);
    }

    public String getDepartment()
    {
        return getString("Department");
    }

    public void setEmail1Address(String emailAddress)
    {
        put("Email1Address", emailAddress);
    }

    public String getEmail1Address()
    {
        return getString("Email1Address");
    }

    public void setEmail1AddressType(String emailAddressType)
    {
        put("Email1AddressType", emailAddressType);
    }

    public String getEmail1AddressType()
    {
        return getString("Email1AddressType");
    }

    public void setEmail1DisplayName(String emailDisplayName)
    {
        put("Email1DisplayName", emailDisplayName);
    }

    public String getEmail1DisplayName()
    {
        return getString("Email1DisplayName");
    }

    public void setEmail1EntryId(String emailEntryId)
    {
        put("Email1EntryID", emailEntryId);
    }

    public String getEmail1EntryId()
    {
        return getString("Email1EntryID");
    }

    public void setEmail2Address(String emailAddress)
    {
        put("Email2Address", emailAddress);
    }

    public String getEmail2Address()
    {
        return getString("Email2Address");
    }

    public void setEmail2AddressType(String emailAddressType)
    {
        put("Email2AddressType", emailAddressType);
    }

    public String getEmail2AddressType()
    {
        return getString("Email2AddressType");
    }

    public void setEmail2DisplayName(String emailDisplayName)
    {
        put("Email2DisplayName", emailDisplayName);
    }

    public String getEmail2DisplayName()
    {
        return getString("Email2DisplayName");
    }

    public void setEmail2EntryId(String emaiEntryId)
    {
        put("Email2EntryID", emaiEntryId);
    }

    public String getEmail2EntryId()
    {
        return getString("Email2EntryID");
    }

    public void setEmail3Address(String emailAddress)
    {
        put("Email3Address", emailAddress);
    }

    public String getEmail3Address()
    {
        return getString("Email3Address");
    }

    public void setEmail3AddressType(String emailAddressType)
    {
        put("Email3AddressType", emailAddressType);
    }

    public String getEmail3AddressType()
    {
        return getString("Email3AddressType");
    }

    public void setEmail3DisplayName(String emailDisplayName)
    {
        put("Email3DisplayName", emailDisplayName);
    }

    public String getEmail3DisplayName()
    {
        return getString("Email3DisplayName");
    }

    public void setEmail3EntryId(String emaiEntryId)
    {
        put("Email3EntryID", emaiEntryId);
    }

    public String getEmail3EntryId()
    {
        return getString("Email3EntryID");
    }

    public void setFileAs(String fileAs)
    {
        put("FileAs", fileAs);
    }

    public String getFileAs()
    {
        return getString("FileAs");
    }

    public void setFirstName(String firstName)
    {
        put("FirstName", firstName);
    }

    public String getFirstName()
    {
        return getString("FirstName");
    }

    public void setFtpSite(String ftpSite)
    {
        put("FTPSite", ftpSite);
    }

    public String getFtpSite()
    {
        return getString("FTPSite");
    }

    public void setFullName(String fullName)
    {
        put("FullName", fullName);
    }

    public String getFullName()
    {
        return getString("FullName");
    }

    public String getFullNameAndCompany()
    {
        return getString("FullNameAndCompany");
    }

    public void setGender(OutlookGender gender)
    {
        put("Gender", gender);
    }

    public OutlookGender getGender()
    {
        return getConstant("Gender", OutlookGender.class);
    }

    public void setGovernmentIdNumber(String governmentIdNumber)
    {
        put("GovernmentIDNumber", governmentIdNumber);
    }

    public String getGovernmentIdNumber()
    {
        return getString("GovernmentIDNumber");
    }

    public boolean hasPicture()
    {
        return getBoolean("HasPicture");
    }

    public void setHobby(String hobby)
    {
        put("Hobby", hobby);
    }

    public String getHobby()
    {
        return getString("Hobby");
    }

    public void setHome2TelephoneNumber(String home2TelephoneNumber)
    {
        put("Home2TelephoneNumber", home2TelephoneNumber);
    }

    public String getHome2TelephoneNumber()
    {
        return getString("Home2TelephoneNumber");
    }

    public void setHomeAddress(String homeAddress)
    {
        put("HomeAddress", homeAddress);
    }

    public String getHomeAddress()
    {
        return getString("HomeAddress");
    }

    public void setHomeAddressCity(String homeAddressCity)
    {
        put("HomeAddressCity", homeAddressCity);
    }

    public String getHomeAddressCity()
    {
        return getString("HomeAddressCity");
    }

    public void setHomeAddressCountry(String homeAddressCountry)
    {
        put("HomeAddressCountry", homeAddressCountry);
    }

    public String getHomeAddressCountry()
    {
        return getString("HomeAddressCountry");
    }

    public void setHomeAddressPostalCode(String homeAddressPostalCode)
    {
        put("HomeAddressPostalCode", homeAddressPostalCode);
    }

    public String getHomeAddressPostalCode()
    {
        return getString("HomeAddressPostalCode");
    }

    public void setHomeAddressPostOfficeBox(String homeAddressPostOfficeBox)
    {
        put("HomeAddressPostOfficeBox", homeAddressPostOfficeBox);
    }

    public String getHomeAddressPostOfficeBox()
    {
        return getString("HomeAddressPostOfficeBox");
    }

    public void setHomeAddressState(String homeAddressState)
    {
        put("HomeAddressState", homeAddressState);
    }

    public String getHomeAddressState()
    {
        return getString("HomeAddressState");
    }

    public void setHomeAddressStreet(String homeAddressStreet)
    {
        put("HomeAddressStreet", homeAddressStreet);
    }

    public String getHomeAddressStreet()
    {
        return getString("HomeAddressStreet");
    }

    public void setHomeFaxNumber(String homeFaxNumber)
    {
        put("HomeFaxNumber", homeFaxNumber);
    }

    public String getHomeFaxNumber()
    {
        return getString("HomeFaxNumber");
    }

    public void setHomeTelephoneNumber(String homeTelephoneNumber)
    {
        put("HomeTelephoneNumber", homeTelephoneNumber);
    }

    public String getHomeTelephoneNumber()
    {
        return getString("HomeTelephoneNumber");
    }

    public void setIMAddress(String imAddress)
    {
        put("IMAddress", imAddress);
    }

    public String getIMAddress()
    {
        return getString("IMAddress");
    }

    public void setInitials(String initials)
    {
        put("Initials", initials);
    }

    public String getInitials()
    {
        return getString("Initials");
    }

    public void setInternetFreeBusyAddress(String internetFreeBusyAddress)
    {
        put("InternetFreeBusyAddress", internetFreeBusyAddress);
    }

    public String getInternetFreeBusyAddress()
    {
        return getString("InternetFreeBusyAddress");
    }

    public void setIsdnNumber(String isdnNumber)
    {
        put("ISDNNumber", isdnNumber);
    }

    public String getIsdnNumber()
    {
        return getString("ISDNNumber");
    }

    public boolean isMarkedAsTask()
    {
        return getBoolean("IsMarkedAsTask");
    }

    public void setJobTitle(String jobTitle)
    {
        put("JobTitle", jobTitle);
    }

    public String getJobTitle()
    {
        return getString("JobTitle");
    }

    public void setJournal(boolean journal)
    {
        put("Journal", journal);
    }

    public boolean isJournal()
    {
        return getBoolean("Journal");
    }

    public void setLanguage(String language)
    {
        put("Language", language);
    }

    public String getLanguage()
    {
        return getString("Language");
    }

    public String getLastFirstAndSuffix()
    {
        return getString("LastFirstAndSuffix");
    }

    public String getLastFirstNoSpace()
    {
        return getString("LastFirstNoSpace");
    }

    public String getLastFirstNoSpaceAndSuffix()
    {
        return getString("LastFirstNoSpaceAndSuffix");
    }

    public String getLastFirstNoSpaceCompany()
    {
        return getString("LastFirstNoSpaceCompany");
    }

    public String getLastFirstSpaceOnly()
    {
        return getString("LastFirstSpaceOnly");
    }

    public String getLastFirstSpaceOnlyCompany()
    {
        return getString("LastFirstSpaceOnlyCompany");
    }

    public void setLastName(String lastName)
    {
        put("LastName", lastName);
    }

    public String getLastName()
    {
        return getString("LastName");
    }

    public String getLastNameAndFirstName()
    {
        return getString("LastNameAndFirstName");
    }

    public void setMailingAddress(String mailingAddress)
    {
        put("MailingAddress", mailingAddress);
    }

    public String getMailingAddress()
    {
        return getString("MailingAddress");
    }

    public void setMailingAddressCity(String mailingAddressCity)
    {
        put("MailingAddressCity", mailingAddressCity);
    }

    public String getMailingAddressCity()
    {
        return getString("MailingAddressCity");
    }

    public void setMailingAddressCountry(String mailingAddressCountry)
    {
        put("MailingAddressCountry", mailingAddressCountry);
    }

    public String getMailingAddressCountry()
    {
        return getString("MailingAddressCountry");
    }

    public void setMailingAddressPostalCode(String mailingAddressPostalCode)
    {
        put("MailingAddressPostalCode", mailingAddressPostalCode);
    }

    public String getMailingAddressPostalCode()
    {
        return getString("MailingAddressPostalCode");
    }

    public void setMailingAddressPostOfficeBox(String mailingAddressPostOfficeBox)
    {
        put("MailingAddressPostOfficeBox", mailingAddressPostOfficeBox);
    }

    public String getMailingAddressPostOfficeBox()
    {
        return getString("MailingAddressPostOfficeBox");
    }

    public void setMailingAddressState(String mailingAddressState)
    {
        put("MailingAddressState", mailingAddressState);
    }

    public String getMailingAddressState()
    {
        return getString("MailingAddressState");
    }

    public void setMailingAddressStreet(String mailingAddressStreet)
    {
        put("MailingAddressStreet", mailingAddressStreet);
    }

    public String getMailingAddressStreet()
    {
        return getString("MailingAddressStreet");
    }

    public void setManagerName(String managerName)
    {
        put("ManagerName", managerName);
    }

    public String getManagerName()
    {
        return getString("ManagerName");
    }

    public void setMiddleName(String middleName)
    {
        put("MiddleName", middleName);
    }

    public String getMiddleName()
    {
        return getString("MiddleName");
    }

    public void setMobileTelephoneNumber(String mobileTelephoneNumber)
    {
        put("MobileTelephoneNumber", mobileTelephoneNumber);
    }

    public String getMobileTelephoneNumber()
    {
        return getString("MobileTelephoneNumber");
    }

    public void setNetMeetingAlias(String netMeetingAlias)
    {
        put("NetMeetingAlias", netMeetingAlias);
    }

    public String getNetMeetingAlias()
    {
        return getString("NetMeetingAlias");
    }

    public void setNetMeetingServer(String netMeetingServer)
    {
        put("NetMeetingServer", netMeetingServer);
    }

    public String getNetMeetingServer()
    {
        return getString("NetMeetingServer");
    }

    public void setNickName(String nickName)
    {
        put("NickName", nickName);
    }

    public String getNickName()
    {
        return getString("NickName");
    }

    public void setOfficeLocation(String officeLocation)
    {
        put("OfficeLocation", officeLocation);
    }

    public String getOfficeLocation()
    {
        return getString("OfficeLocation");
    }

    public void setOrganizationalIdNumber(String organizationalIdNumber)
    {
        put("OrganizationalIDNumber", organizationalIdNumber);
    }

    public String getOrganizationalIdNumber()
    {
        return getString("OrganizationalIDNumber");
    }

    public void setOtherAddress(String otherAddress)
    {
        put("OtherAddress", otherAddress);
    }

    public String getOtherAddress()
    {
        return getString("OtherAddress");
    }

    public void setOtherAddressCity(String otherAddressCity)
    {
        put("OtherAddressCity", otherAddressCity);
    }

    public String getOtherAddressCity()
    {
        return getString("OtherAddressCity");
    }

    public void setOtherAddressCountry(String otherAddressCountry)
    {
        put("OtherAddressCountry", otherAddressCountry);
    }

    public String getOtherAddressCountry()
    {
        return getString("OtherAddressCountry");
    }

    public void setOtherAddressPostalCode(String otherAddressPostalCode)
    {
        put("OtherAddressPostalCode", otherAddressPostalCode);
    }

    public String getOtherAddressPostalCode()
    {
        return getString("OtherAddressPostalCode");
    }

    public void setOtherAddressPostOfficeBox(String otherAddressPostOfficeBox)
    {
        put("OtherAddressPostOfficeBox", otherAddressPostOfficeBox);
    }

    public String getOtherAddressPostOfficeBox()
    {
        return getString("OtherAddressPostOfficeBox");
    }

    public void setOtherAddressState(String otherAddressState)
    {
        put("OtherAddressState", otherAddressState);
    }

    public String getOtherAddressState()
    {
        return getString("OtherAddressState");
    }

    public void setOtherAddressStreet(String otherAddressStreet)
    {
        put("OtherAddressStreet", otherAddressStreet);
    }

    public String getOtherAddressStreet()
    {
        return getString("OtherAddressStreet");
    }

    public void setOtherFaxNumber(String otherFaxNumber)
    {
        put("OtherFaxNumber", otherFaxNumber);
    }

    public String getOtherFaxNumber()
    {
        return getString("OtherFaxNumber");
    }

    public void setOtherTelephoneNumber(String otherTelephoneNumber)
    {
        put("OtherTelephoneNumber", otherTelephoneNumber);
    }

    public String getOtherTelephoneNumber()
    {
        return getString("OtherTelephoneNumber");
    }

    public void setPagerNumber(String pagerNumber)
    {
        put("PagerNumber", pagerNumber);
    }

    public String getPagerNumber()
    {
        return getString("PagerNumber");
    }

    public void setPersonalHomePage(String personalHomePage)
    {
        put("PersonalHomePage", personalHomePage);
    }

    public String getPersonalHomePage()
    {
        return getString("PersonalHomePage");
    }

    public void setPrimaryTelephoneNumber(String primaryTelephoneNumber)
    {
        put("PrimaryTelephoneNumber", primaryTelephoneNumber);
    }

    public String getPrimaryTelephoneNumber()
    {
        return getString("PrimaryTelephoneNumber");
    }

    public void setProfession(String profession)
    {
        put("Profession", profession);
    }

    public String getProfession()
    {
        return getString("Profession");
    }

    public void setRadioTelephoneNumber(String radioTelephoneNumber)
    {
        put("RadioTelephoneNumber", radioTelephoneNumber);
    }

    public String getRadioTelephoneNumber()
    {
        return getString("RadioTelephoneNumber");
    }

    public void setReferredBy(String referredBy)
    {
        put("ReferredBy", referredBy);
    }

    public String getReferredBy()
    {
        return getString("ReferredBy");
    }

    public void setReminderTime(Date reminderTime)
    {
        put("ReminderTime", reminderTime);
    }

    public Date getReminderTime()
    {
        return getDate("ReminderTime");
    }

    public void setRtfBody(String rtfBody)
    {
        put("RTFBody", rtfBody);
    }

    public String getRtfBody()
    {
        return getString("RTFBody");
    }

    public void setSelectedMailingAddress(OutlookMailingAddress selectedMailingAddress)
    {
        put("SelectedMailingAddress", selectedMailingAddress);
    }

    public OutlookMailingAddress getSelectedMailingAddress()
    {
        return getConstant("SelectedMailingAddress", OutlookMailingAddress.class);
    }

    public void setSpouse(String spouse)
    {
        put("Spouse", spouse);
    }

    public String getSpouse()
    {
        return getString("Spouse");
    }

    public void setSuffix(String suffix)
    {
        put("Suffix", suffix);
    }

    public String getSuffix()
    {
        return getString("Suffix");
    }

    public void setTaskCompletedDate(Date taskCompletedDate)
    {
        put("TaskCompletedDate", taskCompletedDate);
    }

    public Date getTaskCompletedDate()
    {
        return getDate("TaskCompletedDate");
    }

    public void setTaskDueDate(Date taskDueDate)
    {
        put("TaskDueDate", taskDueDate);
    }

    public Date getTaskDueDate()
    {
        return getDate("TaskDueDate");
    }

    public void setTaskStartDate(Date taskStartDate)
    {
        put("TaskStartDate", taskStartDate);
    }

    public Date getTaskStartDate()
    {
        return getDate("TaskStartDate");
    }

    public void setTaskSubject(String taskSubject)
    {
        put("TaskSubject", taskSubject);
    }

    public String getTaskSubject()
    {
        return getString("TaskSubject");
    }

    public void setTelexNumber(String telexNumber)
    {
        put("TelexNumber", telexNumber);
    }

    public String getTelexNumber()
    {
        return getString("TelexNumber");
    }

    public void setTitle(String title)
    {
        put("Title", title);
    }

    public String getTitle()
    {
        return getString("Title");
    }

    public void setToDoTaskOrdinal(Date toDoTaskOrdinal)
    {
        put("ToDoTaskOrdinal", toDoTaskOrdinal);
    }

    public Date getToDoTaskOrdinal()
    {
        return getDate("ToDoTaskOrdinal");
    }

    public void setTtytddTelephoneNumber(String ttytddTelephoneNumber)
    {
        put("TTYTDDTelephoneNumber", ttytddTelephoneNumber);
    }

    public String getTtytddTelephoneNumber()
    {
        return getString("TTYTDDTelephoneNumber");
    }

    public void setUserField1(String userField1)
    {
        put("User1", userField1);
    }

    public String getUserField1()
    {
        return getString("User1");
    }

    public void setUserField2(String userField2)
    {
        put("User2", userField2);
    }

    public String getUserField2()
    {
        return getString("User2");
    }

    public void setUserField3(String userField3)
    {
        put("User3", userField3);
    }

    public String getUserField3()
    {
        return getString("User3");
    }

    public void setUserField4(String userField4)
    {
        put("User4", userField4);
    }

    public String getUserField4()
    {
        return getString("User4");
    }

    public void setWebPage(String webPage)
    {
        put("WebPage", webPage);
    }

    public String getWebPage()
    {
        return getString("WebPage");
    }

    public void setYomiCompanyName(String yomiCompanyName)
    {
        put("YomiCompanyName", yomiCompanyName);
    }

    public String getYomiCompanyName()
    {
        return getString("YomiCompanyName");
    }

    public void setYomiFirstName(String yomiFirstName)
    {
        put("YomiFirstName", yomiFirstName);
    }

    public String getYomiFirstName()
    {
        return getString("YomiFirstName");
    }

    public void setYomiLastName(String yomiLastName)
    {
        put("YomiLastName", yomiLastName);
    }

    public String getYomiLastName()
    {
        return getString("YomiLastName");
    }
}
