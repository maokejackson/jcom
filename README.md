 jcom+
=============
Let's automate Microsoft application with Java code.

## Use Cases

Read all contacts in Outlook.
```Java

OutlookApplication outlook = new OutlookApplication();
OutlookFolder folder = outlook.getNameSpace().getFolder(OutlookDefaultFolder.CONTACTS);

for (OutlookItem item : folder.getItems())
{
    OutlookContact contact = item.cast(OutlookContact.class);
    System.out.println(contact.getFirstName());
    System.out.println(contact.getLastName());
    System.out.println(contact.getBirthday());
}
```
Create and send a mail in Outlook.
```Java
OutlookApplication outlook = new OutlookApplication();

OutlookMail mail = outlook.createMail();
mail.setTo("someone@somewhere.net");
mail.setHtmlBody("<html><body>A test mail.</body></html>");
mail.setBodyFormat(OutlookBodyFormat.HTML);
mail.setImportance(OutlookImportance.HIGH);
mail.getAttachments().add("C:\\folder\\filename.txt");

mail.save(); // save mail in draft
mail.send(); // send mail automatically
```
