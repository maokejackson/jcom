package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookDefaultFolderType;
import com.dtxmaker.jcom.outlook.constant.OutlookObjectClass;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

import static org.junit.Assert.*;

public class OutlookApplicationTest
{
    private static OutlookApplication outlook;

    @BeforeClass
    public static void setUp() throws Exception
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

    @AfterClass
    public static void tearDown() throws Exception
    {
        outlook.getNamespace().getDefaultFolder(OutlookDefaultFolderType.DRAFTS).getItems().removeAll();
        outlook.getNamespace().getDefaultFolder(OutlookDefaultFolderType.CONTACTS).getItems().removeAll();
    }

    @Test
    public void testGetObjectClass() throws Exception
    {
        assertEquals(OutlookObjectClass.APPLICATION, outlook.getObjectClass());
    }

    @Test
    public void testGetDefaultProfileName() throws Exception
    {
        assertEquals("Outlook", outlook.getDefaultProfileName());
    }

    @Test
    public void testIsTrusted() throws Exception
    {
        assertFalse(outlook.isTrusted());
    }

    @Test
    public void testGetLanguageSettings() throws Exception
    {
        assertNotNull(outlook.getLanguageSettings());
    }

    @Test
    public void testGetName() throws Exception
    {
        assertEquals("Outlook", outlook.getName());
    }

    @Test
    public void testGetProductCode() throws Exception
    {
        assertEquals("{90120000-0030-0000-0000-0000000FF1CE}", outlook.getProductCode());
    }

    @Test
    public void testGetVersion() throws Exception
    {
        assertEquals("12.0.0.4518", outlook.getVersion());
    }

    @Test
    public void testGetDrafts() throws Exception
    {
        OutlookFolder folder = outlook.getNamespace().getDefaultFolder(OutlookDefaultFolderType.DRAFTS);
        OutlookItem item = folder.getItems().getItem(1);

        assertEquals(OutlookObjectClass.MAIL, item.getObjectClass());
    }

    @Test
    public void testGetContacts() throws Exception
    {
        OutlookFolder folder = outlook.getNamespace().getDefaultFolder(OutlookDefaultFolderType.CONTACTS);
        OutlookItem item = folder.getItems().getItem(1);

        assertEquals(OutlookObjectClass.CONTACT, item.getObjectClass());
    }
}
