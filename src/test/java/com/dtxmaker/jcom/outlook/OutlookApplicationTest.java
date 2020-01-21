package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookDefaultFolder;
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
        mail.attachFile("D:\\tmp\\licenses.xml");
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
        outlook.getFolder(OutlookDefaultFolder.DRAFTS).removeAllItems();
        outlook.getFolder(OutlookDefaultFolder.CONTACTS).removeAllItems();
        outlook.getFolder(OutlookDefaultFolder.DELETED_ITEMS).removeAllItems();
    }

    @Test
    public void testGetDrafts() throws Exception
    {
        OutlookFolder folder = outlook.getFolder(OutlookDefaultFolder.DRAFTS);
        OutlookItem item = folder.getItemAt(1);

        assertTrue(item.isObject(OutlookObjectClass.MAIL));
    }

    @Test
    public void testGetContacts() throws Exception
    {
        OutlookFolder folder = outlook.getFolder(OutlookDefaultFolder.CONTACTS);
        OutlookItem item = folder.getItemAt(1);

        assertTrue(item.isObject(OutlookObjectClass.CONTACT));
    }
}
