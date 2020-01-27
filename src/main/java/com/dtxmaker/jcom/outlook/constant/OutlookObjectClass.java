package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.TypedConstant;
import com.dtxmaker.jcom.outlook.*;
import com.dtxmaker.jcom.util.EnumUtils;

import java.util.Arrays;

/**
 * Specifies constants that represent the different Microsoft Outlook object classes.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/outlook.olobjectclass">
 * https://docs.microsoft.com/en-us/office/vba/api/outlook.olobjectclass</a>
 */
public enum OutlookObjectClass implements TypedConstant<Outlook>
{
    ACCOUNT(105),
    ACCOUNT_RULE_CONDITION(135),
    ACCOUNTS(106),
    ACTION(32),
    ACTIONS(33),
    ADDRESS_ENTRIES(21),
    ADDRESS_ENTRY(8, OutlookAddressEntry.class),
    ADDRESS_LIST(7),
    ADDRESS_LISTS(20),
    ADDRESS_RULE_CONDITION(170),
    APPLICATION(0, OutlookAppointment.class),
    APPOINTMENT(26, OutlookAppointment.class),
    ASSIGN_TO_CATEGORY_RULE_ACTION(122),
    ATTACHMENT(5, OutlookAttachment.class),
    ATTACHMENTS(18, OutlookAttachments.class),
    ATTACHMENT_SELECTION(169),
    AUTO_FORMAT_RULE(147),
    AUTO_FORMAT_RULES(148),
    CALENDAR_MODULE(159),
    CALENDAR_SHARING(151),
    CATEGORIES(153, OutlookCategories.class),
    CATEGORY(152, OutlookCategory.class),
    CATEGORY_RULE_CONDITION(130),
    CLASS_BUSINESS_CARD_VIEW(168),
    CLASS_CALENDAR_VIEW(139),
    CLASS_CARD_VIEW(138),
    CLASS_ICON_VIEW(137),
    CLASS_NAVIGATION_PANE(155),
    CLASS_PEOPLE_VIEW(183),
    CLASS_TABLE_VIEW(136),
    CLASS_TIME_LINE_VIEW(140),
    CLASS_TIME_ZONE(174),
    CLASS_TIME_ZONES(175),
    COLUMN(154),
    COLUMN_FORMAT(149),
    COLUMNS(150),
    CONFLICT(102),
    CONFLICTS(103),
    CONTACT(40, OutlookContact.class),
    CONTACTS_MODULE(160),
    CONVERSATION(178),
    CONVERSATION_HEADER(182),
    DISTRIBUTION_LIST(69),
    DOCUMENT(41),
    EXCEPTION(30),
    EXCEPTIONS(29),
    EXCHANGE_DISTRIBUTION_LIST(111),
    EXCHANGE_USER(110),
    EXPLORER(34),
    EXPLORERS(60),
    FOLDER(2, OutlookFolder.class),
    FOLDERS(15, OutlookFolders.class),
    FOLDER_USER_PROPERTIES(172),
    FOLDER_USER_PROPERTY(171),
    FORM_DESCRIPTION(37),
    FORM_NAME_RULE_CONDITION(131),
    FORM_REGION(129),
    FROM_RSS_FEED_RULE_CONDITION(173),
    FROM_RULE_CONDITION(132),
    IMPORTANCE_RULE_CONDITION(128),
    INSPECTOR(35),
    INSPECTORS(61),
    ITEM_PROPERTIES(98),
    ITEM_PROPERTY(99),
    ITEMS(16, OutlookItems.class),
    JOURNAL(42, OutlookJournal.class),
    JOURNAL_MODULE(162),
    MAIL(43, OutlookMail.class),
    MAIL_MODULE(158),
    MARK_AS_TASK_RULE_ACTION(124),
    MEETING_CANCELLATION(54, OutlookMeeting.class),
    MEETING_FORWARD_NOTIFICATION(181, OutlookMeeting.class),
    MEETING_REQUEST(53, OutlookMeeting.class),
    MEETING_RESPONSE_NEGATIVE(55, OutlookMeeting.class),
    MEETING_RESPONSE_POSITIVE(56, OutlookMeeting.class),
    MEETING_RESPONSE_TENTATIVE(57, OutlookMeeting.class),
    MOVE_OR_COPY_RULE_ACTION(118),
    NAMESPACE(1, OutlookNameSpace.class),
    NAVIGATION_FOLDER(167),
    NAVIGATION_FOLDERS(166),
    NAVIGATION_GROUP(165),
    NAVIGATION_GROUPS(164),
    NAVIGATION_MODULE(157),
    NAVIGATION_MODULES(156),
    NEW_ITEM_ALERT_RULE_ACTION(125),
    NOTE(44, OutlookNote.class),
    NOTES_MODULE(163),
    ORDER_FIELD(144),
    ORDER_FIELDS(145),
    OUTLOOK_BAR_GROUP(66),
    OUTLOOK_BAR_GROUPS(65),
    OUTLOOK_BAR_PANE(63),
    OUTLOOK_BAR_SHORTCUT(68),
    OUTLOOK_BAR_SHORTCUTS(67),
    OUTLOOK_BAR_STORAGE(64),
    OUTSPACE(180),
    PAGES(36),
    PANES(62),
    PLAY_SOUND_RULE_ACTION(123),
    POST(45, OutlookPost.class),
    PROPERTY_ACCESSOR(112),
    PROPERTY_PAGES(71),
    PROPERTY_PAGE_SITE(70),
    RECIPIENT(4, OutlookRecipient.class),
    RECIPIENTS(17, OutlookRecipients.class),
    RECURRENCE_PATTERN(28),
    REMINDER(101),
    REMINDERS(100),
    REMOTE(47),
    REPORT(46),
    RESULTS(78),
    ROW(121),
    RULE(115),
    RULE_ACTION(117),
    RULE_ACTIONS(116),
    RULE_CONDITION(127),
    RULE_CONDITIONS(126),
    RULES(114),
    SEARCH(77),
    SELECTION(74),
    SELECT_NAMES_DIALOG(109),
    SENDER_IN_ADDRESS_LIST_RULE_CONDITION(133),
    SEND_RULE_ACTION(119),
    SHARING(104),
    SIMPLE_ITEMS(179),
    SOLUTIONS_MODULE(177),
    STORAGE_ITEM(113),
    STORE(107),
    STORES(108),
    SYNC_OBJECT(72),
    SYNC_OBJECTS(73),
    TABLE(120),
    TASK(48, OutlookTask.class),
    TASK_REQUEST(49),
    TASK_REQUEST_ACCEPT(51),
    TASK_REQUEST_DECLINE(52),
    TASK_REQUEST_UPDATE(50),
    TASKS_MODULE(161),
    TEXT_RULE_CONDITION(134),
    USER_DEFINED_PROPERTIES(172),
    USER_DEFINED_PROPERTY(171),
    USER_PROPERTIES(38),
    USER_PROPERTY(39),
    VIEW(80),
    VIEW_FIELD(142),
    VIEW_FIELDS(141),
    VIEW_FONT(146),
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
