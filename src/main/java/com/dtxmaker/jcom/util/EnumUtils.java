package com.dtxmaker.jcom.util;

import com.dtxmaker.jcom.Constant;

import java.util.Arrays;

public final class EnumUtils
{
    private EnumUtils()
    {
        throw new Error(getClass() + " contains static methods only");
    }

    public static <E extends Enum<E> & Constant> E findByValue(Class<E> clazz, int value)
    {
        return Arrays.stream(clazz.getEnumConstants())
                .filter(object -> object.getValue() == value)
                .findFirst()
                .orElse(null);
    }
}
