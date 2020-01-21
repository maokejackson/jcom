package com.dtxmaker.jcom.outlook;

import com.dtxmaker.jcom.outlook.constant.OutlookBodyFormat;
import com.dtxmaker.jcom.outlook.constant.OutlookImportance;
import com.jacob.com.Dispatch;

public class OutlookMail extends OutlookItem
{
    public OutlookMail(OutlookApplication application, Dispatch dispatch)
    {
        super(application, dispatch);
    }

    public void setSubject(String subject)
    {
        put("Subject", subject);
    }

    public String getSubject()
    {
        return getString("Subject");
    }

    public void setSenderName(String senderName)
    {
        put("SenderName", senderName);
    }

    public String getSenderName()
    {
        return getString("SenderName");
    }

    public void setTo(String to)
    {
        put("To", to);
    }

    public String getTo()
    {
        return getString("To");
    }

    public void setCc(String cc)
    {
        put("CC", cc);
    }

    public String getCc()
    {
        return getString("CC");
    }

    public void setBcc(String bcc)
    {
        put("BCC", bcc);
    }

    public String getBcc()
    {
        return getString("BCC");
    }

    public void setBody(String body)
    {
        put("Body", body);
    }

    public String getBody()
    {
        return getString("Body");
    }

    public void setHtmlBody(String htmlBody)
    {
        put("HTMLBody", htmlBody);
    }

    public String getHtmlBody()
    {
        return getString("HTMLBody");
    }

    public void setBodyFormat(OutlookBodyFormat bodyFormat)
    {
        put("BodyFormat", bodyFormat.getValue());
    }

    public OutlookBodyFormat getBodyFormat()
    {
        return OutlookBodyFormat.findByValue(getInt("BodyFormat"));
    }

    public void setImportance(OutlookImportance importance)
    {
        put("Importance", importance.getValue());
    }

    public OutlookImportance getImportance()
    {
        return OutlookImportance.findByValue(getInt("Importance"));
    }

    /**
     * Send mail automatically.
     */
    public void send()
    {
        Dispatch.call(dispatch, "Send");
    }
}
