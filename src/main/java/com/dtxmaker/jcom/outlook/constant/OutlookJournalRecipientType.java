package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookJournalRecipientType implements Constant
{
    ASSOCIATED_CONTACT(1),
    ;

    private final int value;

    OutlookJournalRecipientType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookJournalRecipientType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookJournalRecipientType.class, value);
    }
}
