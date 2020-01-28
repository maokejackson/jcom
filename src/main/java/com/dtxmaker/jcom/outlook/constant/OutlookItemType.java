package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.TypedConstant;
import com.dtxmaker.jcom.outlook.*;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookItemType implements TypedConstant<OutlookItem>
{
    MAIL(0, OutlookMail.class),
    APPOINTMENT(1, OutlookAppointment.class),
    CONTACT(2, OutlookContact.class),
    TASK(3, OutlookTask.class),
    JOURNAL(4, OutlookJournal.class),
    NOTE(5, OutlookNote.class),
    POST(6, OutlookPost.class),
    DISTRIBUTION_LIST(7, OutlookDistList.class),
    ;

    private final int                          value;
    private final Class<? extends OutlookItem> type;

    OutlookItemType(int value, Class<? extends OutlookItem> type)
    {
        this.value = value;
        this.type = type;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    @Override
    public Class<? extends OutlookItem> getType()
    {
        return type;
    }

    public static OutlookItemType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookItemType.class, value);
    }

    public static OutlookItemType findByType(Class<? extends OutlookItem> type)
    {
        return EnumUtils.findByType(OutlookItemType.class, type);
    }

    public static int findValueByType(Class<? extends OutlookItem> type)
    {
        return EnumUtils.findValueByType(OutlookItemType.class, type, MAIL);
    }
}
