package com.dtxmaker.jcom.library.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum AppLanguageID implements Constant
{
    APP_LANGUAGE_ID(4),
    HELP(3),
    INSTALL(1),
    UI(2),
    UI_PREVIOUS(5),
    ;

    private final int value;

    AppLanguageID(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static AppLanguageID findByValue(int value)
    {
        return EnumUtils.findByValue(AppLanguageID.class, value);
    }
}
