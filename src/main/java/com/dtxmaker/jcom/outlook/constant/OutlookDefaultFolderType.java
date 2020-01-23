package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookDefaultFolderType implements Constant
{
    DELETED_ITEMS(3),
    OUTBOX(4),
    SENT_MAIL(5),
    INBOX(6),
    CALENDAR(9),
    CONTACTS(10),
    JOURNAL(11),
    NOTES(12),
    TASKS(13),
    DRAFTS(16),
    ALL_PUBLIC_FOLDERS(18),
    CONFLICTS(19),
    SYNC_ISSUES(20),
    LOCAL_FAILURES(21),
    SERVER_FAILURES(22),
    JUNK(23),
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
