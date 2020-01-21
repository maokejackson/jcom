package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookSaveAsType implements Constant
{
    TXT(0),
    RTF(1),
    TEMPLATE(2),
    MSG(3),
    DOC(4),
    HTML(5),
    V_CARD(6),
    V_CAL(7),
    I_CAL(8),
    MSG_UNICODE(9),
    ;

    private final int value;

    OutlookSaveAsType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookSaveAsType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookSaveAsType.class, value);
    }
}
