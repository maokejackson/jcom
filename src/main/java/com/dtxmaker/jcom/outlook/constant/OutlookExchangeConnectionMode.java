package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookExchangeConnectionMode implements Constant
{
    CACHED_CONNECTED_DRIZZLE(600),
    CACHED_CONNECTED_FULL(700),
    CACHED_CONNECTED_HEADERS(500),
    CACHED_DISCONNECTED(400),
    CACHED_OFFLINE(200),
    DISCONNECTED(300),
    NO_EXCHANGE(0),
    ONLINE(800),
    ;

    private final int value;

    OutlookExchangeConnectionMode(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookExchangeConnectionMode findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookExchangeConnectionMode.class, value);
    }
}
