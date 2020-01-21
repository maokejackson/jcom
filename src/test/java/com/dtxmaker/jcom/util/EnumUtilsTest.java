package com.dtxmaker.jcom.util;

import com.dtxmaker.jcom.Constant;
import org.junit.Test;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

public class EnumUtilsTest
{
    @Test
    public void givenStubEnum_WhenValueIsOne_ThenReturnOne() throws Exception
    {
        StubEnum actual = EnumUtils.findByValue(StubEnum.class, 1);

        assertEquals(StubEnum.ONE, actual);
    }

    @Test
    public void givenStubEnum_WhenValueIsTwo_ThenReturnTwo() throws Exception
    {
        StubEnum actual = EnumUtils.findByValue(StubEnum.class, 2);

        assertEquals(StubEnum.TWO, actual);
    }

    @Test
    public void givenStubEnum_WhenValueIsThree_ThenReturnThree() throws Exception
    {
        StubEnum actual = EnumUtils.findByValue(StubEnum.class, 3);

        assertEquals(StubEnum.THREE, actual);
    }

    @Test
    public void givenStubEnum_WhenValueIsFour_ThenReturnNull() throws Exception
    {
        StubEnum actual = EnumUtils.findByValue(StubEnum.class, 4);

        assertNull(actual);
    }

    private enum StubEnum implements Constant
    {
        ONE(1),
        TWO(2),
        THREE(3),
        ;

        private final int value;

        StubEnum(int value)
        {
            this.value = value;
        }

        @Override
        public int getValue()
        {
            return value;
        }
    }
}
