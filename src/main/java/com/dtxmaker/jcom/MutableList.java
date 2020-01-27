package com.dtxmaker.jcom;

public interface MutableList<T> extends ImmutableList<T>
{
    /**
     * Removes an object from the collection.
     *
     * @param index the 1-based index number of the object in the collection.
     */
    void remove(int index);

    /**
     * Remove all objects from the collection.
     */
    void removeAll();
}
