MailUtility:

The main objective of mail utility is to download all the attachments of a mail and place it in local directory.
This utility is developed in such a way that it accepts the 5 arguments from command line.

Step 1:
Place the jar in your local directory and make sure that the directory path is set in environment variables.

Step 2:
As mentioned above utility accepts 5 arguments.
args[0]- email id should be passed (<userName@xyz.com>)
args[1]- email password (<12345>)
args[2]- Folder name should be specified where the email with attachments resides. (Inbox,Drafts,etc)
args[3]- Directory path should be specified where the email attachments should get downloaded (C:\FolderName)
args[4]- Subject pattern should be specified that matches with the initial word in subject of the email (MSI,POS)

Step 3:
User should run the command through command prompt by passing the arguments in the specified order as mentioned below :
Java MailUtility.MailUtility.mailUtilites "email-id" "emailPassword" "InboxfolderName" "DownloadDirectoryPath" "Email Subject Pattern" 