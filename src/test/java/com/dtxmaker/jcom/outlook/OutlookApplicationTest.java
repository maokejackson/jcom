package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookDefaultFolderType;
import com.dtxmaker.jcom.outlook.constant.OutlookObjectClass;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import static org.junit.Assert.assertTrue;

public class OutlookApplicationTest
{
    private OutlookApplication outlook;

    @Before
    public void setUp() throws Exception
    {
        outlook = new OutlookApplication();

        OutlookMail mail = outlook.createMail();
        mail.setTo("maokejackson@gmail.com");
        mail.setSubject("Test Mail");
        mail.setBody("This is a test mail");
        mail.save();

        OutlookContact contact = outlook.createContact();
        contact.setFirstName("Maoke");
        contact.setLastName("Jackson");
        contact.setEmailAddress("maokejackson@gmail.com");
        contact.save();
    }

    @After
    public void tearDown() throws Exception
    {
        outlook.getDefaultFolder(OutlookDefaultFolderType.DRAFTS).removeAllItems();
        outlook.getDefaultFolder(OutlookDefaultFolderType.CONTACTS).removeAllItems();
        outlook.getDefaultFolder(OutlookDefaultFolderType.DELETED_ITEMS).removeAllItems();
    }

    @Test
    public void testGetDrafts() throws Exception
    {
        OutlookDefaultFolder folder = outlook.getDefaultFolder(OutlookDefaultFolderType.DRAFTS);
        OutlookItem item = folder.getItemAt(1);

        assertTrue(item.isObject(OutlookObjectClass.MAIL));
    }

    @Test
    public void testGetContacts() throws Exception
    {
        OutlookDefaultFolder folder = outlook.getDefaultFolder(OutlookDefaultFolderType.CONTACTS);
        OutlookItem item = folder.getItemAt(1);

        assertTrue(item.isObject(OutlookObjectClass.CONTACT));
    }
}
