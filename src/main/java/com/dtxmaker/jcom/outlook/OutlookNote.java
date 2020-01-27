package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

/**
 * Represents a note in a Notes folder.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.noteitem">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.noteitem</a>
 */
public class OutlookNote extends OutlookItem
{
    OutlookNote(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    /**
     * Sets the height (in pixels) of the note window.
     *
     * @param height the height
     */
    public void setHeight(int height)
    {
        put("Height", height);
    }

    /**
     * Returns the height (in pixels) of the note window.
     */
    public int getHeight()
    {
        return getInt("Height");
    }

    /**
     * Sets the position (in pixels) of the left vertical edge of a note window from the edge of the screen.
     *
     * @param left the left position
     */
    public void setLeft(int left)
    {
        put("Left", left);
    }

    /**
     * Returns the position (in pixels) of the left vertical edge of a note window from the edge of the screen.
     */
    public int getLeft()
    {
        return getInt("Left");
    }

    /**
     * Sets the position (in pixels) of the top horizontal edge of a note window from the edge of the screen.
     *
     * @param top the top position
     */
    public void setTop(int top)
    {
        put("Top", top);
    }

    /**
     * Returns the position (in pixels) of the top horizontal edge of a note window from the edge of the screen.
     */
    public int getTop()
    {
        return getInt("Top");
    }

    /**
     * Sets the width (in pixels) of the specified object.
     *
     * @param width the width
     */
    public void setWidth(int width)
    {
        put("Width", width);
    }

    /**
     * Returns the width (in pixels) of the specified object.
     */
    public int getWidth()
    {
        return getInt("Width");
    }
}
