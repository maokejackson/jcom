package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.TypedConstant;
import com.dtxmaker.jcom.outlook.*;
import com.dtxmaker.jcom.util.EnumUtils;

import java.util.Arrays;

public enum OutlookObjectClass implements TypedConstant<Outlook>
{
    ACTION(32),
    ACTIONS(33),
    ADDRESS_ENTRIES(21),
    ADDRESS_ENTRY(8, OutlookAddressEntry.class),
    ADDRESS_LIST(7),
    ADDRESS_LISTS(20),
    APPLICATION(0, OutlookAppointment.class),
    APPOINTMENT(26, OutlookAppointment.class),
    ATTACHMENT(5, OutlookAttachment.class),
    ATTACHMENTS(18, OutlookAttachments.class),
    CONFLICT(11),
    CONFLICTS(11),
    CONTACT(40, OutlookContact.class),
    DISTRIBUTION_LIST(69),
    DOCUMENT(41),
    EXCEPTION(30),
    EXCEPTIONS(29),
    EXPLORER(34),
    EXPLORERS(60),
    FOLDER(2, OutlookFolder.class),
    FOLDERS(15, OutlookFolders.class),
    FORM_DESCRIPTION(37),
    INSPECTOR(35),
    INSPECTORS(61),
    ITEM_PROPERTIES(98),
    ITEM_PROPERTY(99),
    ITEMS(16, OutlookItems.class),
    JOURNAL(42, OutlookJournal.class),
    LINK(75),
    LINKS(76),
    MAIL(43, OutlookMail.class),
    MEETING_CANCELLATION(54),
    MEETING_REQUEST(53),
    MEETING_RESPONSE_NEGATIVE(55),
    MEETING_RESPONSE_POSITIVE(56),
    MEETING_RESPONSE_TENTATIVE(57),
    NAMESPACE(1, OutlookNameSpace.class),
    NOTE(44, OutlookNote.class),
    OUTLOOK_BAR_GROUP(66),
    OUTLOOK_BAR_GROUPS(65),
    OUTLOOK_BAR_PANE(63),
    OUTLOOK_BAR_SHORTCUT(68),
    OUTLOOK_BAR_SHORTCUTS(67),
    OUTLOOK_BAR_STORAGE(64),
    PAGES(36),
    PANES(62),
    POST(45, OutlookPost.class),
    PROPERTY_PAGES(71),
    PROPERTY_PAGE_SITE(70),
    RECIPIENT(4, OutlookRecipient.class),
    RECIPIENTS(17, OutlookRecipients.class),
    RECURRENCE_PATTERN(28),
    REMINDER(10),
    REMINDERS(10),
    REMOTE(47),
    REPORT(46),
    RESULTS(78),
    SEARCH(77),
    SELECTION(74),
    SYNC_OBJECT(72),
    SYNC_OBJECTS(73),
    TASK(48, OutlookTask.class),
    TASK_REQUEST(49),
    TASK_REQUEST_ACCEPT(51),
    TASK_REQUEST_DECLINE(52),
    TASK_REQUEST_UPDATE(50),
    USER_PROPERTIES(38),
    USER_PROPERTY(39),
    VIEW(80),
    VIEWS(79),
    ;

    private final int                      value;
    private final Class<? extends Outlook> type;

    OutlookObjectClass(int value)
    {
        this(value, Outlook.class);
    }

    OutlookObjectClass(int value, Class<? extends Outlook> type)
    {
        this.value = value;
        this.type = type;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    @Override
    public Class<? extends Outlook> getType()
    {
        return type;
    }

    public static OutlookObjectClass findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookObjectClass.class, value);
    }

    public static OutlookObjectClass findByType(Class<? extends Outlook> type)
    {
        return Arrays.stream(values())
                .filter(object -> object.getType() == type)
                .findFirst()
                .orElse(null);
    }
}
