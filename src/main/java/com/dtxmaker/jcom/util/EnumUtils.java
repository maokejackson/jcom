package com.dtxmaker.jcom.util;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.TypedConstant;

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

    public static <T, E extends Enum<E> & TypedConstant<T>> E findByType(Class<E> clazz, Class<? extends T> type)
    {
        return Arrays.stream(clazz.getEnumConstants())
                .filter(object -> object.getType() == type)
                .findFirst()
                .orElse(null);
    }

    public static <T, E extends Enum<E> & TypedConstant<T>> int findValueByType(Class<E> clazz, Class<? extends T> type,
            E defaultValue)
    {
        return Arrays.stream(clazz.getEnumConstants())
                .filter(object -> object.getType() == type)
                .findFirst()
                .orElse(defaultValue)
                .getValue();
    }
}
