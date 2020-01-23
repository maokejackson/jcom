package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookDefaultFolderType;
import org.junit.After;
import org.junit.Before;
import org.junit.Ignore;
import org.junit.Test;

import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

public class OutlookFolderTest
{
    private OutlookApplication outlook;

    @Before
    public void setUp() throws Exception
    {
        outlook = new OutlookApplication();

        OutlookDefaultFolder inbox = outlook.getDefaultFolder(OutlookDefaultFolderType.INBOX);
        inbox.addFolder("Foo");
        inbox.addFolder("Bar");

        OutlookMail mail = outlook.createMail();
        mail.setTo("maokejackson@gmail.com");
        mail.attachFile("D:\\tmp\\licenses.xml");
        mail.save();
    }

    @After
    public void tearDown() throws Exception
    {
        outlook.getDefaultFolder(OutlookDefaultFolderType.INBOX).removeAllFolders();
        outlook.getDefaultFolder(OutlookDefaultFolderType.DRAFTS).removeAllItems();
        outlook.getDefaultFolder(OutlookDefaultFolderType.DELETED_ITEMS).removeAllItems();
    }

    @Test
    public void testGetFolder() throws Exception
    {
        OutlookDefaultFolder inbox = outlook.getDefaultFolder(OutlookDefaultFolderType.INBOX);
        OutlookFolder folder = inbox.getFolder("Foo");

        assertEquals("Foo", folder.getName());
    }

    @Test
    public void testGetFolderAt() throws Exception
    {
        OutlookDefaultFolder inbox = outlook.getDefaultFolder(OutlookDefaultFolderType.INBOX);
        OutlookFolder folder = inbox.getFolderAt(2);

        assertEquals("Bar", folder.getName());
    }

    @Ignore
    @Test
    public void testSetName() throws Exception
    {
        OutlookDefaultFolder inbox = outlook.getDefaultFolder(OutlookDefaultFolderType.INBOX);
        OutlookFolder folder = inbox.getFolder("Foo");
        folder.setName("Foo1");

        assertEquals("Foo1", folder.getName());
    }

    @Test
    public void testGetFolderCount() throws Exception
    {
        OutlookDefaultFolder folder = outlook.getDefaultFolder(OutlookDefaultFolderType.INBOX);

        assertEquals(2, folder.getFolderCount());
    }

    @Test
    public void testGetItemCount() throws Exception
    {
        OutlookDefaultFolder folder = outlook.getDefaultFolder(OutlookDefaultFolderType.DRAFTS);

        assertEquals(1, folder.getItemCount());
    }

    @Test
    public void testGetItems() throws Exception
    {
        OutlookDefaultFolder folder = outlook.getDefaultFolder(OutlookDefaultFolderType.DRAFTS);
        List<OutlookMail> mails = folder.getItems(OutlookMail.class);

        assertNotNull(mails);
        assertEquals(1, mails.size());
    }
}
