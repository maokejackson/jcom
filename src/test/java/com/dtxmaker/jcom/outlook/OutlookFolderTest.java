package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookDefaultFolderType;
import org.junit.After;
import org.junit.Before;
import org.junit.Ignore;
import org.junit.Test;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

public class OutlookFolderTest
{
    private OutlookApplication outlook;

    @Before
    public void setUp()
    {
        outlook = new OutlookApplication();

        OutlookFolder inbox = outlook.getNamespace().getDefaultFolder(OutlookDefaultFolderType.INBOX);
        inbox.getFolders().add("Foo");
        inbox.getFolders().add("Bar");

        OutlookMail mail = outlook.createMail();
        mail.setTo("maokejackson@gmail.com");
        mail.setSubject("Test Mail");
        mail.setBody("This is a test mail");
        mail.save();
    }

    @After
    public void tearDown()
    {
        outlook.getNamespace().getDefaultFolder(OutlookDefaultFolderType.INBOX).getFolders().removeAll();
        outlook.getNamespace().getDefaultFolder(OutlookDefaultFolderType.DRAFTS).getItems().removeAll();
        outlook.getNamespace().getDefaultFolder(OutlookDefaultFolderType.DELETED_ITEMS).getFolders().removeAll();
        outlook.getNamespace().getDefaultFolder(OutlookDefaultFolderType.DELETED_ITEMS).getItems().removeAll();
    }

    @Test
    public void testGetFolder()
    {
        OutlookFolder inbox = outlook.getNamespace().getDefaultFolder(OutlookDefaultFolderType.INBOX);
        OutlookFolder folder = inbox.getFolder("Foo");

        assertEquals("Foo", folder.getName());
    }

    @Test
    public void testGetFolderAt()
    {
        OutlookFolder inbox = outlook.getNamespace().getDefaultFolder(OutlookDefaultFolderType.INBOX);
        OutlookFolder folder = inbox.getFolders().getItem(2);

        assertEquals("Bar", folder.getName());
    }

    @Ignore
    @Test
    public void testSetName()
    {
        OutlookFolder inbox = outlook.getNamespace().getDefaultFolder(OutlookDefaultFolderType.INBOX);
        OutlookFolder folder = inbox.getFolder("Foo");
        folder.setName("Foo1");

        assertEquals("Foo1", folder.getName());
    }

    @Test
    public void testGetFolderCount()
    {
        OutlookFolder folder = outlook.getNamespace().getDefaultFolder(OutlookDefaultFolderType.INBOX);

        assertEquals(2, folder.getFolders().getCount());
    }

    @Test
    public void testGetItemCount()
    {
        OutlookFolder folder = outlook.getNamespace().getDefaultFolder(OutlookDefaultFolderType.DRAFTS);

        assertEquals(1, folder.getItems().getCount());
    }

    @Test
    public void testGetFirst()
    {
        OutlookFolder deletedItems = outlook.getNamespace().getDefaultFolder(OutlookDefaultFolderType.DELETED_ITEMS);

        assertNull(deletedItems.getFolders().getFirst());
    }
}
