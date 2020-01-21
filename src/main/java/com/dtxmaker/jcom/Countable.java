package com.dtxmaker.jcom;

public interface Countable<E>
{
    /**
     * Return the number of items.
     *
     * @return the number of items.
     */
    int getItemCount();

    /**
     * Return an array of items. Zero size array is returned when no items.
     *
     * @return an array of items.
     */
    E[] getItems();

    /**
     * Return the item at specific <code>index</code>.
     *
     * @param index index of item to return, start from 1, order by date ascending.
     * @return the item at specific <code>index</code>.
     */
    E getItemAt(int index);

    /**
     * Permanently remove all items.
     */
    void removeAllItems();

    /**
     * Permanently remove item at specific <code>index</code>.
     *
     * @param index the index of item, start from 1.
     */
    void removeItemAt(int index);
}
