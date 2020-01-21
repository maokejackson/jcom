 jcom+
=============
Let's automate Microsoft application with Java code.

## Use Cases

Read all contacts in Outlook.
```Java

OutlookApplication outlook = new OutlookApplication();
OutlookFolder folder = outlook.getFolder(OutlookDefaultFolder.CONTACTS);
List<OutlookContact> contacts = folder.getItems(OutlookContact.class);

for (OutlookContact contact : contacts)
{
    System.out.println(contact.getFirstName());
    System.out.println(contact.getLastName());
    System.out.println(contact.getBirthday());
    System.out.println(contact.getNotes());
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
mail.attachFile("D:\\tmp\\licenses.xml");

mail.save(); // save mail in draft
mail.send(); // send mail automatically
```
