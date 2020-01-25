package com.dtxmaker.jcom.library.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum AppLanguageId implements Constant
{
    APP_LANGUAGE_ID(4),
    HELP(3),
    INSTALL(1),
    UI(2),
    UI_PREVIOUS(5),
    ;

    private final int value;

    AppLanguageId(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static AppLanguageId findByValue(int value)
    {
        return EnumUtils.findByValue(AppLanguageId.class, value);
    }
}
