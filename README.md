# OL-Save_Attachments
VBA Script to save attachments to local storage and remove them from emails to save space in the mailbox

## Still has a few bugs!

Doesn't save / remove all attchments when there are more than one on a message. [See issue 1](https://github.com/dostergaard/OL-Save_Attachments/issues/1)

## Description
This little VBA script will iterate through a set of selected messages saving each attachment to a folder corresponding to the conversation topic, subject, or "Uncategorized". It will delete the attachments from the mail message and insert links into the message indicating where the attachment has been saved.

Attachments are stored in folder under a folder named "EmailAttachments" in your documents folder. If you want to store them someplace else change the line of code below.

The code will ignore attachments starting with "image" which are typically embedded images such as signature logos or images pasted directly into the message body.

To use it, select one or more messages in a mail or search folder then run "SaveAllAttachments". You will be prompted to allow or deny access to modify messages. You can choose to allow for a period of time or you will be prompted for each attachment.
