package com.dtxmaker.jcom;

public interface MutableList<T, V> extends ImmutableList<T>
{
    /**
     * Creates a new object in the collection.
     *
     * @param value the value for the new object
     * @return a new object.
     */
    T add(V value);

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
