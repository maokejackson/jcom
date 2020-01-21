package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookOfficeDocItemsType implements Constant
{
    EXCEL_WORK_SHEET(8),
    WORD_DOCUMENT(9),
    POWER_POINT_SHOW(10),
    ;

    private final int value;

    OutlookOfficeDocItemsType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookOfficeDocItemsType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookOfficeDocItemsType.class, value);
    }
}
