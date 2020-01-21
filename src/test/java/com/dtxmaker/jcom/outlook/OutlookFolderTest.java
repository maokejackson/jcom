package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookDefaultFolder;
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

        OutlookFolder inbox = outlook.getFolder(OutlookDefaultFolder.INBOX);
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
        outlook.getFolder(OutlookDefaultFolder.INBOX).removeAllFolders();
        outlook.getFolder(OutlookDefaultFolder.DRAFTS).removeAllItems();
        outlook.getFolder(OutlookDefaultFolder.DELETED_ITEMS).removeAllItems();
    }

    @Test
    public void testGetFolder() throws Exception
    {
        OutlookFolder inbox = outlook.getFolder(OutlookDefaultFolder.INBOX);
        OutlookFolder folder = inbox.getFolder("Foo");

        assertEquals("Foo", folder.getName());
    }

    @Test
    public void testGetFolderAt() throws Exception
    {
        OutlookFolder inbox = outlook.getFolder(OutlookDefaultFolder.INBOX);
        OutlookFolder folder = inbox.getFolderAt(2);

        assertEquals("Bar", folder.getName());
    }

    @Ignore
    @Test
    public void testSetName() throws Exception
    {
        OutlookFolder inbox = outlook.getFolder(OutlookDefaultFolder.INBOX);
        OutlookFolder folder = inbox.getFolder("Foo");
        folder.setName("Foo1");

        assertEquals("Foo1", folder.getName());
    }

    @Test
    public void testGetFolderCount() throws Exception
    {
        OutlookFolder folder = outlook.getFolder(OutlookDefaultFolder.INBOX);

        assertEquals(2, folder.getFolderCount());
    }

    @Test
    public void testGetItemCount() throws Exception
    {
        OutlookFolder folder = outlook.getFolder(OutlookDefaultFolder.DRAFTS);

        assertEquals(1, folder.getItemCount());
    }

    @Test
    public void testGetItems() throws Exception
    {
        OutlookFolder folder = outlook.getFolder(OutlookDefaultFolder.DRAFTS);
        List<OutlookMail> mails = folder.getItems(OutlookMail.class);

        assertNotNull(mails);
        assertEquals(1, mails.size());
    }
}
