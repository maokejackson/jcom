package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookAddressEntryUserType implements Constant
{
    EXCHANGE_AGENT(3),
    EXCHANGE_DISTRIBUTION_LIST(1),
    EXCHANGE_ORGANIZATION(4),
    EXCHANGE_PUBLIC_FOLDER(2),
    EXCHANGE_REMOTE_USER(5),
    EXCHANGE_USER(0),
    LDAP(20),
    OTHER(40),
    OUTLOOK_CONTACT(10),
    OUTLOOK_DISTRIBUTION_LIST(11),
    SMTP(30),
    ;

    private final int value;

    OutlookAddressEntryUserType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookAddressEntryUserType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookAddressEntryUserType.class, value);
    }
}
