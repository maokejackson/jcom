package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookBusyStatus;
import com.dtxmaker.jcom.outlook.constant.OutlookMeetingStatus;
import com.dtxmaker.jcom.outlook.constant.OutlookRecurrenceState;
import com.dtxmaker.jcom.outlook.constant.OutlookResponseStatus;
import com.jacob.com.Dispatch;

import java.util.Date;

/**
 * Represents a meeting, a one-time appointment, or a recurring appointment or meeting in the Calendar folder.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.appointmentitem">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.appointmentitem</a>
 */
public class OutlookAppointment extends AbstractOutlookInternalItem
{
    OutlookAppointment(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Methods                        *
     *                                                     *
     *******************************************************/

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    /**
     * Sets if the appointment is an all-day event (as opposed to a specified time).
     *
     * @param allDayEvent <code>true</code> if the appointment is an all-day event
     */
    public void setAllDayEvent(boolean allDayEvent)
    {
        put("AllDayEvent", allDayEvent);
    }

    /**
     * Returns <code>true</code> if the appointment is an all-day event (as opposed to a specified time).
     */
    public boolean isAllDayEvent()
    {
        return getBoolean("AllDayEvent");
    }

    /**
     * Returns the busy status of the user for the appointment.
     *
     * @param busyStatus the busy status
     */
    public void setBusyStatus(OutlookBusyStatus busyStatus)
    {
        put("BusyStatus", busyStatus);
    }

    /**
     * Returns the busy status of the user for the appointment.
     */
    public OutlookBusyStatus getBusyStatus()
    {
        return getConstant("BusyStatus", OutlookBusyStatus.class);
    }

    /**
     * Sets the duration (in minutes) of the AppointmentItem.
     *
     * @param duration the duration (in minutes)
     */
    public void setDuration(int duration)
    {
        put("Duration", duration);
    }

    /**
     * Returns the duration (in minutes) of the AppointmentItem.
     */
    public int getDuration()
    {
        return getInt("Duration");
    }

    /**
     * Sets the end date and time of a Appointment entry.
     *
     * @param end the end date and time
     */
    public void setEnd(Date end)
    {
        put("End", end);
    }

    /**
     * Returns the end date and time of a Appointment entry.
     */
    public Date getEnd()
    {
        return getDate("End");
    }

    /**
     * Sets the end date and time of the appointment expressed in the {@link #getEndTimeZone()}.
     *
     * @param endInEndTimeZone the end date and time
     */
    public void setEndInEndTimeZone(Date endInEndTimeZone)
    {
        put("EndInEndTimeZone", endInEndTimeZone);
    }

    /**
     * Returns the end date and time of the appointment expressed in the {@link #getEndTimeZone()}.
     */
    public Date getEndInEndTimeZone()
    {
        return getDate("EndInEndTimeZone");
    }

    /**
     * Sets a TimeZone value that corresponds to the end time of the appointment.
     *
     * @param endTimeZone a TimeZone value that corresponds to the end time of the appointment
     */
    public void setEndTimeZone(OutlookTimeZone endTimeZone)
    {
        put("EndTimeZone", endTimeZone);
    }

    /**
     * Returns a TimeZone value that corresponds to the end time of the appointment.
     */
    public OutlookTimeZone getEndTimeZone()
    {
        return new OutlookTimeZone(getDispatch("EndTimeZone"));
    }

    /**
     * Sets the end date and time of the appointment expressed in the Coordinated Universal Time (UTC) standard.
     *
     * @param endUtc the end date and time of the appointment
     */
    public void setEndUtc(Date endUtc)
    {
        put("EndUTC", endUtc);
    }

    /**
     * Returns the end date and time of the appointment expressed in the Coordinated Universal Time (UTC) standard.
     */
    public Date getEndUtc()
    {
        return getDate("EndUTC");
    }

    /**
     * Sets whether updates to the AppointmentItem object should be sent to all attendees.
     *
     * @param forceUpdateToAllAttendees <code>true</code> to sent to all attendees when appointment is updated
     */
    public void setForceUpdateToAllAttendees(boolean forceUpdateToAllAttendees)
    {
        put("ForceUpdateToAllAttendees", forceUpdateToAllAttendees);
    }

    /**
     * Indicates whether updates to the AppointmentItem object should be sent to all attendees.
     */
    public boolean isForceUpdateToAllAttendees()
    {
        return getBoolean("ForceUpdateToAllAttendees");
    }

    /**
     * Returns a unique global identifier for the AppointmentItem object.
     */
    public String getGlobalAppointmentID()
    {
        return getString("GlobalAppointmentID");
    }

    /**
     * Sets the Internet code page used by the item.
     *
     * @param internetCodePage the Internet code page
     */
    public void setInternetCodePage(long internetCodePage)
    {
        put("InternetCodepage", internetCodePage);
    }

    /**
     * Returns the Internet code page used by the item.
     */
    public long getInternetCodePage()
    {
        return getLong("InternetCodepage");
    }

    /**
     * Indicates if the appointment is a recurring appointment.
     */
    public boolean isRecurring()
    {
        return getBoolean("IsRecurring");
    }

    /**
     * Sets the specific office location (for example, Building 1 Room 1 or Suite 123) for the appointment.
     *
     * @param location the specific office location
     */
    public void setLocation(String location)
    {
        put("Location", location);
    }

    /**
     * Returns the specific office location (for example, Building 1 Room 1 or Suite 123) for the appointment.
     */
    public String getLocation()
    {
        return getString("Location");
    }

    /**
     * Sets the meeting status of the appointment.
     *
     * @param meetingStatus the meeting status
     */
    public void setMeetingStatus(OutlookMeetingStatus meetingStatus)
    {
        put("MeetingStatus", meetingStatus);
    }

    /**
     * Returns the meeting status of the appointment.
     */
    public OutlookMeetingStatus getMeetingStatus()
    {
        return getConstant("MeetingStatus", OutlookMeetingStatus.class);
    }

    /**
     * Returns the URL for the Meeting Workspace that the appointment item is linked to.
     */
    public String getMeetingWorkspaceURL()
    {
        return getString("MeetingWorkspaceURL");
    }

    /**
     * Sets the display string of optional attendees names for the appointment.
     *
     * @param optionalAttendees the display string of optional attendees names
     */
    public void setOptionalAttendees(String optionalAttendees)
    {
        put("OptionalAttendees", optionalAttendees);
    }

    /**
     * Returns the display string of optional attendees names for the appointment.
     */
    public String getOptionalAttendees()
    {
        return getString("OptionalAttendees");
    }

    /**
     * Returns the name of the organizer of the appointment.
     */
    public String getOrganizer()
    {
        return getString("Organizer");
    }

    /**
     * Returns all the recipients for the Outlook item.
     */
    public final OutlookRecipients getRecipients()
    {
        return new OutlookRecipients(callDispatch("Recipients"));
    }

    /**
     * Returns the recurrence property of the specified object.
     */
    public OutlookRecurrenceState getRecurrenceState()
    {
        return getConstant("RecurrenceState", OutlookRecurrenceState.class);
    }

    /**
     * Sets the number of minutes the reminder should occur prior to the start of the appointment.
     *
     * @param reminderMinutesBeforeStart the number of minutes
     */
    public void setReminderMinutesBeforeStart(int reminderMinutesBeforeStart)
    {
        put("ReminderMinutesBeforeStart", reminderMinutesBeforeStart);
    }

    /**
     * Returns the number of minutes the reminder should occur prior to the start of the appointment.
     */
    public int getReminderMinutesBeforeStart()
    {
        return getInt("ReminderMinutesBeforeStart");
    }

    /**
     * Sets if the reminder overrides the default reminder behavior for the item.
     *
     * @param reminderOverrideDefault <code>true</code> if the reminder overrides the default reminder.
     */
    public void setReminderOverrideDefault(boolean reminderOverrideDefault)
    {
        put("ReminderOverrideDefault", reminderOverrideDefault);
    }

    /**
     * Returns <code>true</code> if the reminder overrides the default reminder behavior for the item.
     */
    public boolean isReminderOverrideDefault()
    {
        return getBoolean("ReminderOverrideDefault");
    }

    /**
     * Sets if the reminder should play a sound when it occurs for this item.
     *
     * @param reminderPlaySound <code>true</code> if the reminder should play a sound.
     */
    public final void setReminderPlaySound(boolean reminderPlaySound)
    {
        put("ReminderPlaySound", reminderPlaySound);
    }

    /**
     * Returns <code>true</code> if the reminder should play a sound when it occurs for this item.
     */
    public final boolean isReminderPlaySound()
    {
        return getBoolean("ReminderPlaySound");
    }

    /**
     * Sets if a reminder has been set for this item.
     *
     * @param reminderSet <code>true</code> if a reminder has been set.
     */
    public final void setReminderSet(boolean reminderSet)
    {
        put("ReminderSet", reminderSet);
    }

    /**
     * Returns <code>true</code> if a reminder has been set for this item.
     */
    public final boolean getReminderSet()
    {
        return getBoolean("ReminderSet");
    }

    /**
     * Sets the path and file name of the sound file to play when the reminder occurs for the Outlook item.
     *
     * @param reminderSoundFile the path and file name of the sound file.
     */
    public final void setReminderSoundFile(String reminderSoundFile)
    {
        put("ReminderSoundFile", reminderSoundFile);
    }

    /**
     * Returns the path and file name of the sound file to play when the reminder occurs for the Outlook item.
     */
    public final String getReminderSoundFile()
    {
        return getString("ReminderSoundFile");
    }

    /**
     * Sets the reply time for the appointment.
     *
     * @param replyTime the reply time
     */
    public void setReplyTime(Date replyTime)
    {
        put("ReplyTime", replyTime);
    }

    /**
     * Returns the reply time for the appointment.
     */
    public Date getReplyTime()
    {
        return getDate("ReplyTime");
    }

    /**
     * Sets required attendee names for the meeting appointment.
     *
     * @param requiredAttendees semicolon-delimited names
     */
    public void setRequiredAttendees(String requiredAttendees)
    {
        put("RequiredAttendees", requiredAttendees);
    }

    /**
     * Returns required attendee names (semicolon-delimited) for the meeting appointment.
     */
    public String getRequiredAttendees()
    {
        return getString("RequiredAttendees");
    }

    /**
     * Sets resource names for the meeting.
     *
     * @param resources semicolon-delimited names
     */
    public void setResources(String resources)
    {
        put("Resources", resources);
    }

    /**
     * Returns resource names (semicolon-delimited) for the meeting.
     */
    public String getResources()
    {
        return getString("Resources");
    }

    /**
     * Sets whether the sender would like a response to the meeting request for the appointment.
     *
     * @param responseRequested <code>true</code> if the sender would like a response
     */
    public void setResponseRequested(boolean responseRequested)
    {
        put("ResponseRequested", responseRequested);
    }

    /**
     * Indicates whether the sender would like a response to the meeting request for the appointment.
     */
    public boolean isResponseRequested()
    {
        return getBoolean("ResponseRequested");
    }

    /**
     * Returns the overall status of the meeting for the current user for the appointment.
     */
    public OutlookResponseStatus getResponseStatus()
    {
        return getConstant("ResponseStatus", OutlookResponseStatus.class);
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
     * Sets the start date and time of the appointment expressed in the {@link #getEndTimeZone()}.
     *
     * @param startInStartTimeZone the start date and time
     */
    public void setStartInStartTimeZone(Date startInStartTimeZone)
    {
        put("StartInStartTimeZone", startInStartTimeZone);
    }

    /**
     * Returns the start date and time of the appointment expressed in the {@link #getStartTimeZone()}.
     */
    public Date getStartInStartTimeZone()
    {
        return getDate("StartInStartTimeZone");
    }

    /**
     * Sets a TimeZone value that corresponds to the time zone for the start time of the appointment.
     *
     * @param startTimeZone a TimeZone value that corresponds to the time zone for the start time of the appointment.
     */
    public void setStartTimeZone(OutlookTimeZone startTimeZone)
    {
        put("StartTimeZone", startTimeZone);
    }

    /**
     * Returns a TimeZone value that corresponds to the time zone for the start time of the appointment.
     */
    public OutlookTimeZone getStartTimeZone()
    {
        return new OutlookTimeZone(getDispatch("StartTimeZone"));
    }

    /**
     * Sets the start date and time of the appointment expressed in the Coordinated Universal Time (UTC) standard.
     *
     * @param startUtc the start date and time of the appointment
     */
    public void setStartUtc(Date startUtc)
    {
        put("StartUTC", startUtc);
    }

    /**
     * Returns the start date and time of the appointment expressed in the Coordinated Universal Time (UTC) standard.
     */
    public Date getStartUtc()
    {
        return getDate("StartUTC");
    }
}
