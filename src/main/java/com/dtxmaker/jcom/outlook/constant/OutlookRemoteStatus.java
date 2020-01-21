package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookRemoteStatus implements Constant
{
    REMOTE_STATUS_NONE(0),
    UN_MARKED(1),
    MARKED_FOR_DOWNLOAD(2),
    MARKED_FOR_COPY(3),
    MARKED_FOR_DELETE(4),
    ;

    private final int value;

    OutlookRemoteStatus(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookRemoteStatus findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookRemoteStatus.class, value);
    }
}
