# Classic ASP Class for PostMarkApp #

Postmark enables web applications of any size to deliver and track transactional email reliably, with minimal setup time and zero maintenance.
http://www.postmarkapp.com/

**17/06/2011 - v1 - Initial Release**

This class allows you to use the Postmark API within a Classic ASP web-application.

Admittedly, it's not feature complete (i.e. it doesn't cover every single feature available via the Postmark API) but it was developed to be used as part of a large-scale mailing application and performs perfectly in that instance.

Please feel free to submit suggestions, changes & bug fixes.


---

# Example Usage #

## Step 1 ##
Modify the Postmark API key (POSTMARK\_API\_KEY) in the Postmark.asp file to be your key.
If you don't do this, you won't be able to send mail.

**NOTE: The POSTMARK\_API\_TESTMODE constant is set to True by default.**

  * True:  Use for testing - you can send to Postmark and recieve successful response, but will NOT actually send the e-mail to recipient.

  * False: Use when going live - this will send to Postmark which will then send e-mails.

## Step 2 ##
Include **postmark.asp** in your code. It will automatically include **json2.asp**, which is required to interact with the Postmark API.
```
<!--#include file="postmark.asp" -->
```

## Step 3 ##
Use the following example code below to start working with Postmark.

There are a couple of functions to add multiple recipients, CC's and BCC's.
```
  SetTo: Single recipient
  SetToCC: Carbon Copy - See single recipient
  SetToBCC: Blind Carbon Copy - Set single recipient
  AddTo: Multiple recipients
  AddToCC: Carbon Copy - Add another recipient
  AddToBC: Blind Carbon Copy - Add another recipient
```
### Code Sample ###
```
Dim PostMarkEmail: Set PostMarkEmail = new PostMark

PostMarkEmail.SetTo("to@address.com") 
PostMarkEmail.AddTo("to-another@address.com") 
PostMarkEmail.SetFrom("from@address.com")
PostMarkEmail.SetSubject("Subject goes here")
' Plain text
PostMarkEmail.SetTextBody("Body of e-mail goes here")
' HTML content
PostMarkEmail.SetHTMLBody("<html><body><h1>Body of email goes here.</h1></body></html>")
PostMarkEmail.Send()

If (PostMarkEmail.SendSuccessful()) Then
  response.write "E-mail was sent!<br />"
  response.write PostMarkEmail.GetMessageID &"<br />"
Else
  response.write "E-mail failed to send...<br />"
  response.write PostMarkEmail.GetErrorCode &"<br />"
  response.write PostMarkEmail.GetMessage &"<br />"
End If

Set PostMarkEmail = Nothing
```