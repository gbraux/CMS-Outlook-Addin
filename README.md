# CMS-Outlook-Add-In
Guillaume BRAUX (gubraux@cisco.com) - Paris, France.

A one-click Add-In to Outlook 2013, 2016 & Web/Mobile to add your Cisco Meeting Server personal meeting room details to your Outlook meeting request body. It also handles room-based video endpoints reservation & "One-Button-To-Push" (Cisco TMS-XE mandatory) if those rooms are added as participants in the meeting request.

**Note : This is a proof-of-concept, developed and tested only into a lab environment. Trying to implement it as-is into a production envt may not be possible, or may require advanced tweaking at code-level**

![My image](https://raw.githubusercontent.com/gbraux/CMS-Outlook-Addin/master/BookingAddin1-edit.png)
![My image](https://raw.githubusercontent.com/gbraux/CMS-Outlook-Addin/master/BookingAddin2-edit.png)

# Features
-	Server Side Add-In, nothing to install on the outlook clients (Pushed by exchange server to all the clients).
-	(Should) works on Outlook 2013 & 2016, and also Outlook Web Access and Outlook Mobile apps. Currently only tested with Outlook 2016 Heavy Client.
-	Meeting details automatically grabbed from CMS (using CMS API) – it uses the default CMS "coSpace" of the Outlook logged user. It includes meeting URI, WebRTC link, PIN …
-	The meeting details are integrated into the meeting invite body, like the "Webex Productivity Tools"
-	Calendar location field is also updated with basic meeting details
-	One Button To Push (OBTP) support on Cisco Video endpoints if TMS-XE is also installed

# Techical details

## Add-In Location

The Outlook Add-In itself is located in the OutlookAdd-In folder. Those files are automatically downloaded by the client when the Add-In is pushed by the Exchange Server (note : not all files may be needed for the Add-In to work, just using the default template from MSFT ...).

The most important files are :

- CMS_Add-In_Manifest.xml : The descriptor of the Add-In, where to get necessary files through HTTP, etc ... This needs to be configured. This is the file that you have to load into Exchange Server when installing the Add-In.

- FunctionFile/Function.js : The most important file, as it is the core logic of the Add-In (ie. what happens when you click the Add-In button in Outlook !)

## Server-Side scripts

Because of limitations of "light" javascript-based Outlook Add-ins (supported since Outlook 2013 and pushed by the Exchange Server), we also need 2 server-side PHP scripts hosted on a web server to handle some stuff the client-side Add-In can't do itself (think about API calls to CMS, or some Exchange Web Services calls).
Those scripts HAVE to be hosted on the same Web Server as the Add-In (no support for CrossDomain calls), and are called by the Add-In through Javascript AJAX calls.

- CmsProxy.php : Server side PHP script to make REST requests to the CMS Server API on behalf on the Add-In, and get the default space details of a user
- EwsProxy.php : Server side PHP script to set the UCCapabilities property of a calendar Item through EWS (Exchange Web Services) on behalf of the Add-In.

## OBTP Support (AKA "The-Ugly-Hack")

This Outlook Add-In mimics the way CMR-Cloud OBTP works by provisioning the same calendar custom property ("UCCapabilies") as the Webex PTools (so TMS-XE "thinks" it is a CMR-Cloud meeting, and schedules a TMS ExternalBridge). This property is provisioned by a call to EWS (Exchange Web Services) through the EwsProxy.php script (the client-side Add-In has no capability to edit the custom property that the PTools are creating - or at least not easily).

BUT ... it is not so simple ... The issue is that, if you programmatically add the required property to the draft calendar item through EWS (you need to programaticaly "save"  the meeting request on Outlook before so the draft can by synced to Exchange and found by EWS) ... this property will be overwritten (or deleted) when the user sends the meeting request (as the local version of the message on Outlook is considered as the most recent one, and all server-side changes on the draft are lost ...)

So, to be persistent, the custom property have to be created right AFTER the meeting request is sent ... But the Outlook JS API does not have any event to know when he message is sent !
"The-Ugly-Hack" is the following : when the add-in button is pressed, it calls the remote EwsProxy.php script, and do not wait for any answer. The server-side script will then run, and loop on the calendar properties every 5 seconds (!) until MeetingRequestWasSent = true. It then writes the custom "UCCapabilies" property.

The loop maximum execution time is limited by timeouts of the Web-Server and PHP. Default is 5 min in my lab environment. Which means that ... if you click the add-in ... and take MORE than 5 minute to set your meeting request details and send it ... the "UCCapabilies" will never be set ... and TMS-XE will book its default bridge, and not an ExternalBridge.

There may be something to do with EWS notification, but I didn't had time to dig into it ...

## User identification

This Add-In uses the default AD-Synced coSpace of the current outlook user as the meeting point. It gets Outlook User Identity (email), and uses the CMS API user search capability (?filter=) to find the corresponding user on CMS (and grab coSpace details).
If the user CMS URI is NOT the email address, some codes have to be tweaked to send the right search string to CMS so it can find the right user (ie. if email is gubraux@cisco.com, but CMS user URI is gubraux.cms@cisco.com ...).

## EWS Impersonation

The add-in (through server-side PHP Scripts) have to make calls to EWS (Echanges Web Services) for advanced features (ie. set OBTP / UCCapabilites property to the calendar item). Impersonation is used so the script (through EWS) can get access to the current mailbox of the user making the request. As such, a super user (with impersonation enabled) needs to be used to make the EWS calls.

# Install

1. Copy all files to a HTTPS + PHP enabled Web Server (note : Outlook seems to check for SSL certificate, so ensure that your web server certificate is trusted by the clients)
2. Edit Config.php file with necessary informations
3. Configure the Add-In XML manifest (CMS_Add-In_Manifest.xml) with the right HTTPS domain/paths of your web server
4. Upload the Add-In XML manifest to your exchange server Add-Ins repository. For testing, Add-Ins can be configured at user/mailbox level from the Heavy Client (Outlook > Account Infomation > Manage Add-Ins) or from Outlook Web Access (Gear icon > Manage Apps). Users may need specific Exchange permission to be able to install add-ins themselves.
