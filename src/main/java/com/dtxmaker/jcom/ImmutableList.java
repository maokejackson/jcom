package com.dtxmaker.jcom;

public interface ImmutableList<T>
{
    /**
     * Returns an object from the collection.
     *
     * @param index the 1-based index number of the object in the collection.
     * @return the specified object.
     */
    T getItem(int index);

    /**
     * Returns the count of objects in the specified collection.
     */
    int getCount();
}
