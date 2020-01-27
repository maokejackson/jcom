package com.dtxmaker.jcom;

public interface TypedConstant<T> extends Constant
{
    Class<? extends T> getType();
}
