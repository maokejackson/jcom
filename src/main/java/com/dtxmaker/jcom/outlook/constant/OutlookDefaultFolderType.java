package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookDefaultFolderType implements Constant
{
    ALL_PUBLIC_FOLDERS(18),
    CALENDAR(9),
    CONTACTS(10),
    CONFLICTS(19),
    DELETED_ITEMS(3),
    DRAFTS(16),
    INBOX(6),
    JOURNAL(11),
    JUNK(23),
    LOCAL_FAILURES(21),
    MANAGED_EMAIL(29),
    NOTES(12),
    OUTBOX(4),
    SENT_MAIL(5),
    SERVER_FAILURES(22),
    SUGGESTED_CONTACTS(30),
    SYNC_ISSUES(20),
    TASKS(13),
    TO_DO(28),
    RSS_FEEDS(25),
    ;

    private final int value;

    OutlookDefaultFolderType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookDefaultFolderType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookDefaultFolderType.class, value);
    }
}
