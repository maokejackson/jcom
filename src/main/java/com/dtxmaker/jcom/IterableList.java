package com.dtxmaker.jcom;

public interface IterableList<T> extends Iterable<T>
{
    /**
     * Returns the first object in the collection.
     *
     * @return the first object contained by the collection.
     */
    T getFirst();

    /**
     * Returns the last object in the collection.
     *
     * @return the last object contained by the collection.
     */
    T getLast();

    /**
     * Returns the next object in the collection.
     *
     * @return the next object contained by the collection.
     */
    T getNext();

    /**
     * Returns the previous object in the collection.
     *
     * @return the previous object contained by the collection.
     */
    T getPrevious();
}
