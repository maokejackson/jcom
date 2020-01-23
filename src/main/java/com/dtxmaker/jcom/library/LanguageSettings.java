package com.dtxmaker.jcom.library;

import com.dtxmaker.jcom.library.constant.AppLanguageID;
import com.dtxmaker.jcom.library.constant.LanguageID;
import com.jacob.com.Dispatch;

public class LanguageSettings
{
    private final Dispatch dispatch;

    public LanguageSettings(Dispatch dispatch) {this.dispatch = dispatch;}

    public int getCreator()
    {
        return Dispatch.get(dispatch, "Creator").getInt();
    }

    public int getLanguageID(AppLanguageID id)
    {
        return Dispatch.call(dispatch, "LanguageID", id.getValue()).getInt();
    }

    public boolean isLanguagePreferredForEditing(LanguageID languageID)
    {
        return Dispatch.call(dispatch, "LanguagePreferredForEditing", languageID.getValue()).getBoolean();
    }
}
