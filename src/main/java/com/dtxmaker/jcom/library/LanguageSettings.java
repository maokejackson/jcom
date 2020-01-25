package com.dtxmaker.jcom.library;

import com.dtxmaker.jcom.Base;
import com.dtxmaker.jcom.library.constant.AppLanguageId;
import com.dtxmaker.jcom.library.constant.LanguageId;
import com.jacob.com.Dispatch;

/**
 * Returns information about the language settings in a Microsoft Office application.
 *
 * @see <a href="https://docs.microsoft.com/en-us/office/vba/api/office.languagesettings">
 * https://docs.microsoft.com/en-us/office/vba/api/office.languagesettings</a>
 */
public class LanguageSettings extends Base
{
    public LanguageSettings(Dispatch dispatch)
    {
        super(dispatch);
    }

    /* *****************************************************
     *                                                     *
     *                      Properties                     *
     *                                                     *
     *******************************************************/

    /**
     * Returns an integer that indicates the application in which the LanguageSettings object was created.
     */
    public int getCreator()
    {
        return getInt("Creator");
    }

    /**
     * Returns an enumeration representing the locale identifier (LCID) for the install language, the user interface language, or the Help language.
     *
     * @param id One of the {@link AppLanguageId} enumerations.
     */
    public int getLanguageID(AppLanguageId id)
    {
        return callInt("LanguageID", id);
    }

    /**
     * Returns <code>true</code> if the value for the LanguageID constant has been identified in the Windows registry as a preferred language for editing.
     *
     * @param languageId One of the {@link LanguageId} enumerations.
     * @return <code>true</code> if the value for the LanguageID constant has been identified in the Windows registry as a preferred language for editing.
     */
    public boolean isLanguagePreferredForEditing(LanguageId languageId)
    {
        return callBoolean("LanguagePreferredForEditing", languageId);
    }
}
