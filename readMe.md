#Using MSGraph to email from Outlook365

The purpose of this project is to provide a demonstration and a template to be able to interact with an Outlook365 mailbox using MSGraph. This will include...

- Accessing a mailbox, whether your own or someone else's
- gather and interact with the messages
	- read a message
	- extract attachments to file
	- filter messages and attachments
	- mark a message read
	- move a message to a folder
- send messages
	- create a new message
	- create a draft of a replay
	- forward a message
	- attach a file to the message
	- add one or more recipients to emails

This is, by no means, a comprehensive list. There is plenty more you can do. The [API documenation](https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview?view=graph-rest-1.0) covers that.

## Requisites

A basic understanding of MSGraph is required. Please start with the MS Graph [Quick Start](https://developer.microsoft.com/en-us/graph/quick-start). This will instruct you on how to register an app for testing. You will need data from that, or another test app registration, to use this code.
Based on the code I wrote, you will need several NuGet libraries. I offer a list of them.
You will also access to one Outlook365 email account. This may be an account you login to, or an alternate account you wish to access. That account should have one email with a safe attachment. You also will need several email addresses for recipients. A safe file different from the one attached to the email would be handy.

## Files
- Program.cs - the main example code
- settings.cs - a class which reads settings from the JSON file
- appsettings.json - the settings file, containing the app registration details
- nuget_install.ps1 - the listing of needed NuGet packages, written as a script you can use to load them into your project
- readMe.md - this document

## App Registration - API Permissions
To use any of these interactions, you will need the correct API permissions on your app registration, based on how you wish to deploy the app using.

- Mail.Read is needed to read mail lists or individual messages.
- Mail.ReadWrite is needed if you also wish to mark as read or move to folder.
- Mail.Send is needed if you wish to send a message, new, reply or forward.
- Append .Shared to any of these if you wish to interact with mailboxes other than one for the logged-in user on a Delegated app. 

you can test permissions using the [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer).

## Warning
Using code to access an Outlook365 mailbox comes with risks and responsibilities. Please keep security in mind. Make sure interactions comply with organizational policies, Microsoft Terms of Services, and some common sense and respect for others.
The code offered is meant only to be an example of some of what is possible, and is offered as-is. While this code is based on what I have used, and has worked for me, I cannot guarantee it will work for you as-is. 

## Back Story
I was caught out when Microsoft closed  basic authentication to Exchange Web Services on 2022 October, and needed to find an alternative fast. I chose MSGraph, since it seemed like the 'simpler' option. It also seemed like the option that has the potential for longer support. I had a time trying to find how I could programatically interact with emails for my apps running in Windows Scheduler. Thus, now that I have my ducklings in a formation, I wish to present a reference and a starting point by which others can more simply wrap their heads around this tool. It is not a perfect tool. It has its issues. It has potential. In a competent and mindful head, it can be put to good and safe use. 
I am not a professional programmer, and I do not play one on Netflix. I am the IT guy for a small business who knows enough programming to be dangerous. My code likely reflects this. 