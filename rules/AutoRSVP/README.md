# Description

This file is a macro module export for Outlook 2016 and up

It contains macros that can be used in an Outlook rule to automatically respond to meeting requests (Accepting or Declining)

# How to use

* Enable Developer Mode: See [Enables Developer Ribbon in Outlook](../../settings#enables-developer-ribbon-in-outlook)
* Enable Unsafe Client Mail rules: See [Enables Developer Ribbon in Outlook](../../settings#enables-unsafe-rules)
* Set Macro Security Level to Notification Only: See [Set Macro Security Level to Notification Only](../../settings#set-macro-security-level-to-notification-only)
* Download AutoRSVP Module: [AutoRSVP.bas](../../../../raw/main/rules/AutoRSVP/AutoRSVP.bas)
* Import BAS Module file to Outlook:
  > `Ribbon` > `Developer` > `Visual Basic` > `(New Window)` > `File` > `Import`
* Add a new Outlook Rule for incomming messages
  * For the Actions of the rule: Select **only**
    * `run a script`
    * `stop procesing more rules`
  * Configure the `run a script` rule and select one of the AutoRSVP macro
    Macro | Action
    ----- | ------
    AcceptAndSend | Accept the meeting and send a response to the organizer
    AcceptSilently | Accept the meeting, but do not send a response to the organizer
    DeclineAndSend | Decline the meeting and send a response to the organizer
    DeclineSilently | Decline the meeting, but do not send a response to the organizer

* Your rule should be configured and working, you can verify what meetings are accepted or declined by:
  * Going to your calendar
  * Checking the Sent folder and see what responses were sent
  * Checking the Deleted folder and see what meetings were deleted
    * In case you use one of the `Silently` maro, you will see the meeting response in the deleted folder as well.
  
