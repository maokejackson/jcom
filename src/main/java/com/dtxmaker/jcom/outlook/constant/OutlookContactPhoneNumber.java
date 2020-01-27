package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookContactPhoneNumber implements Constant
{
    ASSISTANT(0),
    BUSINESS(1),
    BUSINESS_2(2),
    BUSINESS_FAX(3),
    CALLBACK(4),
    CAR(5),
    COMPANY(6),
    HOME(7),
    HOME_2(8),
    HOME_FAX(9),
    ISDN(10),
    MOBILE(11),
    OTHER(12),
    OTHER_FAX(13),
    PAGER(14),
    PRIMARY(15),
    RADIO(16),
    TELEX(17),
    TTYTTD(18),
    ;

    private final int value;

    OutlookContactPhoneNumber(int value) {this.value = value;}

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookContactPhoneNumber findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookContactPhoneNumber.class, value);
    }
}
