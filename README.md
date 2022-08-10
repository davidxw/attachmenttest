# Graph API Message Attachment Test
Demonstrate using Graph API to read mail messages and attachments. To run this demo:

* Create an application in AAD, no specific permissions required
* Add the tenant id and client id as config items (either create a appsettings.json file, or pass in as paramaters)
* Run code

The code will pull the top 20 (unless overridden in config) mail messages that contain attachments from the users inbox, and download the attachments. It will then repeat the process with the attachment downloads occuring in parallel.
