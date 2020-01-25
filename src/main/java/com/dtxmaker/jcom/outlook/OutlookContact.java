package com.dtxmaker.jcom.outlook;

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

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    public String getFirstName()
    {
        return getString("FirstName");
    }

    public void setFirstName(String firstName)
    {
        put("FirstName", firstName);
    }

    public String getMiddleName()
    {
        return getString("MiddleName");
    }

    public void setMiddleName(String middleName)
    {
        put("MiddleName", middleName);
    }

    public String getLastName()
    {
        return getString("LastName");
    }

    public void setLastName(String lastName)
    {
        put("LastName", lastName);
    }

    public String getNickName()
    {
        return getString("NickName");
    }

    public void setNickName(String nickName)
    {
        put("NickName", nickName);
    }

    public String getTitle()
    {
        return getString("Title");
    }

    public void setTitle(String title)
    {
        put("Title", title);
    }

    public String getSuffix()
    {
        return getString("Suffix");
    }

    public void setSuffix(String suffix)
    {
        put("Suffix", suffix);
    }

    public Date getBirthday()
    {
        return getDate("Birthday");
    }

    public void setBirthday(String birthday)
    {
        put("Birthday", birthday);
    }

    public void setBirthday(Date birthday)
    {
        put("Birthday", birthday);
    }

    public Date getAnniversary()
    {
        return getDate("Anniversary");
    }

    public void setAnniversary(String anniversary)
    {
        put("Anniversary", anniversary);
    }

    public void setAnniversary(Date anniversary)
    {
        put("Anniversary", anniversary);
    }

    public String getBusinessAddressStreet()
    {
        return getString("BusinessAddressStreet");
    }

    public void setBusinessAddressStreet(String businessAddressStreet)
    {
        put("BusinessAddressStreet", businessAddressStreet);
    }

    public String getHomeAddressStreet()
    {
        return getString("HomeAddressStreet");
    }

    public void setHomeAddressStreet(String homeAddressStreet)
    {
        put("HomeAddressStreet", homeAddressStreet);
    }

    public String getOtherAddressStreet()
    {
        return getString("OtherAddressStreet");
    }

    public void setOtherAddressStreet(String otherAddressStreet)
    {
        put("OtherAddressStreet", otherAddressStreet);
    }

    public String getMobileTelephoneNumber()
    {
        return getString("MobileTelephoneNumber");
    }

    public void setMobileTelephoneNumber(String mobileTelephoneNumber)
    {
        put("MobileTelephoneNumber", mobileTelephoneNumber);
    }

    public String getBusinessTelephoneNumber()
    {
        return getString("BusinessTelephoneNumber");
    }

    public void setBusinessTelephoneNumber(String businessTelephoneNumber)
    {
        put("BusinessTelephoneNumber", businessTelephoneNumber);
    }

    public String getBusiness2TelephoneNumber()
    {
        return getString("Business2TelephoneNumber");
    }

    public void setBusiness2TelephoneNumber(String business2TelephoneNumber)
    {
        put("Business2TelephoneNumber", business2TelephoneNumber);
    }

    public String getHomeTelephoneNumber()
    {
        return getString("HomeTelephoneNumber");
    }

    public void setHomeTelephoneNumber(String homeTelephoneNumber)
    {
        put("HomeTelephoneNumber", homeTelephoneNumber);
    }

    public String getHome2TelephoneNumber()
    {
        return getString("Home2TelephoneNumber");
    }

    public void setHome2TelephoneNumber(String home2TelephoneNumber)
    {
        put("Home2TelephoneNumber", home2TelephoneNumber);
    }

    public String getHomeFaxNumber()
    {
        return getString("HomeFaxNumber");
    }

    public void setHomeFaxNumber(String homeFaxNumber)
    {
        put("HomeFaxNumber", homeFaxNumber);
    }

    public String getAssistantTelephoneNumber()
    {
        return getString("AssistantTelephoneNumber");
    }

    public void setAssistantTelephoneNumber(String assistantTelephoneNumber)
    {
        put("AssistantTelephoneNumber", assistantTelephoneNumber);
    }

    public String getCallbackTelephoneNumber()
    {
        return getString("CallbackTelephoneNumber");
    }

    public void setCallbackTelephoneNumber(String callbackTelephoneNumber)
    {
        put("CallbackTelephoneNumber", callbackTelephoneNumber);
    }

    public String getCarTelephoneNumber()
    {
        return getString("CarTelephoneNumber");
    }

    public void setCarTelephoneNumber(String carTelephoneNumber)
    {
        put("CarTelephoneNumber", carTelephoneNumber);
    }

    public String getCompanyTelephoneNumber()
    {
        return getString("CompanyMainTelephoneNumber");
    }

    public void setCompanyTelephoneNumber(String companyTelephoneNumber)
    {
        put("CompanyMainTelephoneNumber", companyTelephoneNumber);
    }

    public String getEmailAddress()
    {
        return getString("Email1Address");
    }

    public void setEmailAddress(String emailAddress)
    {
        put("Email1Address", emailAddress);
    }

    public String getEmail2Address()
    {
        return getString("Email2Address");
    }

    public void setEmail2Address(String email2Address)
    {
        put("Email2Address", email2Address);
    }

    public String getEmail3Address()
    {
        return getString("Email3Address");
    }

    public void setEmail3Address(String email3Address)
    {
        put("Email3Address", email3Address);
    }

    public String getBusinessFaxNumber()
    {
        return getString("BusinessFaxNumber");
    }

    public void setBusinessFaxNumber(String businessFaxNumber)
    {
        put("BusinessFaxNumber", businessFaxNumber);
    }

    public String getIsdnNumber()
    {
        return getString("ISDNNumber");
    }

    public void setIsdnNumber(String isdnNumber)
    {
        put("ISDNNumber", isdnNumber);
    }

    public String getOtherTelephoneNumber()
    {
        return getString("OtherTelephoneNumber");
    }

    public void setOtherTelephoneNumber(String otherTelephoneNumber)
    {
        put("OtherTelephoneNumber", otherTelephoneNumber);
    }

    public String getOtherFaxNumber()
    {
        return getString("OtherFaxNumber");
    }

    public void setOtherFaxNumber(String otherFaxNumber)
    {
        put("OtherFaxNumber", otherFaxNumber);
    }

    public String getPagerNumber()
    {
        return getString("PagerNumber");
    }

    public void setPagerNumber(String pagerNumber)
    {
        put("PagerNumber", pagerNumber);
    }

    public String getPrimaryTelephoneNumber()
    {
        return getString("PrimaryTelephoneNumber");
    }

    public void setPrimaryTelephoneNumber(String primaryTelephoneNumber)
    {
        put("PrimaryTelephoneNumber", primaryTelephoneNumber);
    }

    public String getRadioTelephoneNumber()
    {
        return getString("RadioTelephoneNumber");
    }

    public void setRadioTelephoneNumber(String radioTelephoneNumber)
    {
        put("RadioTelephoneNumber", radioTelephoneNumber);
    }

    public String getTelexNumber()
    {
        return getString("TelexNumber");
    }

    public void setTelexNumber(String telexNumber)
    {
        put("TelexNumber", telexNumber);
    }

    public String getTtytddTelephoneNumber()
    {
        return getString("TTYTDDTelephoneNumber");
    }

    public void setTtytddTelephoneNumber(String ttytddTelephoneNumber)
    {
        put("TTYTDDTelephoneNumber", ttytddTelephoneNumber);
    }

    public String getCompanyName()
    {
        return getString("CompanyName");
    }

    public void setCompanyName(String companyName)
    {
        put("CompanyName", companyName);
    }

    public String getJobTitle()
    {
        return getString("JobTitle");
    }

    public void setJobTitle(String jobTitle)
    {
        put("JobTitle", jobTitle);
    }

    public String getDepartment()
    {
        return getString("Department");
    }

    public void setDepartment(String department)
    {
        put("Department", department);
    }

    public String getProfession()
    {
        return getString("Profession");
    }

    public void setProfession(String profession)
    {
        put("Profession", profession);
    }

    public String getSpouse()
    {
        return getString("Spouse");
    }

    public void setSpouse(String spouse)
    {
        put("Spouse", spouse);
    }

    public int getSelectedMailingAddress()
    {
        return getInt("SelectedMailingAddress");
    }

    public void setSelectedMailingAddress(int selectedMailingAddress)
    {
        put("SelectedMailingAddress", selectedMailingAddress);
    }

    public String getWebPage()
    {
        return getString("WebPage");
    }

    public void setWebPage(String webPage)
    {
        put("WebPage", webPage);
    }

    public String getNotes()
    {
        return getString("Body");
    }

    public void setNotes(String notes)
    {
        put("Body", notes);
    }

    public String getUserField1()
    {
        return getString("User1");
    }

    public void setUserField1(String userField1)
    {
        put("User1", userField1);
    }

    public String getUserField2()
    {
        return getString("User2");
    }

    public void setUserField2(String userField2)
    {
        put("User2", userField2);
    }

    public String getUserField3()
    {
        return getString("User3");
    }

    public void setUserField3(String userField3)
    {
        put("User3", userField3);
    }

    public String getUserField4()
    {
        return getString("User4");
    }

    public void setUserField4(String userField4)
    {
        put("User4", userField4);
    }
}
