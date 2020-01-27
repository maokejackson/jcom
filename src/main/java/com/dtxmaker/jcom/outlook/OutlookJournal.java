package com.dtxmaker.jcom.outlook;

import com.jacob.com.Dispatch;

import java.util.Date;

/**
 * Represents a journal entry in a Journal folder.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.journalitem">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.journalitem</a>
 */
public class OutlookJournal extends AbstractOutlookInternalItem
{
    OutlookJournal(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /**
     * Executes the Forward action for an item and returns the resulting copy.
     *
     * @return the new mail item.
     */
    public OutlookJournal forward()
    {
        return new OutlookJournal(callDispatch("Forward"));
    }

    /**
     * Creates a reply, pre-addressed to the original sender, from the original message.
     *
     * @return A MailItem object that represents the reply.
     */
    public OutlookMail reply()
    {
        return new OutlookMail(callDispatch("Reply"));
    }

    /**
     * Creates a reply to all original recipients from the original message.
     *
     * @return A MailItem object that represents the reply.
     */
    public OutlookMail replyAll()
    {
        return new OutlookMail(callDispatch("ReplyAll"));
    }

    /**
     * Starts the timer on the journal entry.
     */
    public void startTimer()
    {
        call("StartTimer");
    }

    /**
     * Stops the timer on the journal entry.
     */
    public void stopTimer()
    {
        call("StopTimer");
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    /**
     * Sets the contact names associated with the Outlook item.
     *
     * @param contactNames the contact names
     */
    public void setContactNames(String contactNames)
    {
        put("ContactNames", contactNames);
    }

    /**
     * Returns the contact names associated with the Outlook item.
     */
    public String getContactNames()
    {
        return getString("ContactNames");
    }

    /**
     * Sets whether the journalized item was posted as part of the journalized session.
     *
     * @param docPosted <code>true</code> if the journalized item was posted
     */
    public void setDocPosted(boolean docPosted)
    {
        put("DocPosted", docPosted);
    }

    /**
     * Indicates whether the journalized item was posted as part of the journalized session.
     */
    public boolean isDocPosted()
    {
        return getBoolean("DocPosted");
    }

    /**
     * Sets whether the journalized item was printed as part of the journalized session.
     *
     * @param docPrinted <code>true</code> if the journalized item was printed
     */
    public void setDocPrinted(boolean docPrinted)
    {
        put("DocPrinted", docPrinted);
    }

    /**
     * Indicates whether the journalized item was printed as part of the journalized session.
     */
    public boolean isDocPrinted()
    {
        return getBoolean("DocPrinted");
    }

    /**
     * Sets whether the journalized item was routed as part of the journalized session.
     *
     * @param docRouted <code>true</code> if the journalized item was routed
     */
    public void setDocRouted(boolean docRouted)
    {
        put("DocRouted", docRouted);
    }

    /**
     * Indicates whether the journalized item was routed as part of the journalized session.
     */
    public boolean isDocRouted()
    {
        return getBoolean("DocRouted");
    }

    /**
     * Sets whether the journalized item was saved as part of the journalized session.
     *
     * @param docSaved <code>true</code> if the journalized item was saved
     */
    public void setDocSaved(boolean docSaved)
    {
        put("DocSaved", docSaved);
    }

    /**
     * Indicates whether the journalized item was saved as part of the journalized session.
     */
    public boolean isDocSaved()
    {
        return getBoolean("DocSaved");
    }

    /**
     * Sets the duration (in minutes) of the JournalItem.
     *
     * @param duration the duration (in minutes)
     */
    public void setDuration(int duration)
    {
        put("Duration", duration);
    }

    /**
     * Returns the duration (in minutes) of the JournalItem.
     */
    public int getDuration()
    {
        return getInt("Duration");
    }

    /**
     * Sets the end date and time of a Journal entry.
     *
     * @param end the end date and time
     */
    public void setEnd(Date end)
    {
        put("End", end);
    }

    /**
     * Returns the end date and time of a Journal entry.
     */
    public Date getEnd()
    {
        return getDate("End");
    }

    /**
     * Returns all the recipients for the Outlook item.
     */
    public final OutlookRecipients getRecipients()
    {
        return new OutlookRecipients(callDispatch("Recipients"));
    }

    /**
     * Sets the starting date and time for the Outlook item.
     *
     * @param start the starting date and time
     */
    public void setStart(Date start)
    {
        put("Start", start);
    }

    /**
     * Returns the starting date and time for the Outlook item.
     */
    public Date getStart()
    {
        return getDate("Start");
    }

    /**
     * Sets a free-form field, usually containing the display name of the journalizing application (for example, "MSWord".)
     *
     * @param type the display name of the journalizing application
     */
    public void setType(String type)
    {
        put("Type", type);
    }

    /**
     * Returns a free-form field, usually containing the display name of the journalizing application (for example, "MSWord".)
     */
    public String getType()
    {
        return getString("Type");
    }
}
