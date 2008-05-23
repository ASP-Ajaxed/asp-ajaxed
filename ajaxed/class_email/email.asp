<%
'**************************************************************************************************************

'' @CLASSTITLE:		Email
'' @CREATOR:		Michal Gabrukiewicz - gabru @ grafix.at
'' @CREATEDON:		23.04.2008
'' @CDESCRIPTION:	This class represents an email. It is a generic interface which internally uses
''					a third party component.
''					- create a new instance, set the properties and send or user newWith(template) and create a new instance with a template which will set the body and the subject of the email automatically from the TextTemplate
''					- you should send an email whereever you want to send an email in your application. Even if you dont know if an email component will be installed. 
'' @REQUIRES:		TextTemplate
'' @COMPATIBLE:		DIMAC JMail.Message, ASPEMAIL Persists.MailSender, MS cdo.message
'' @VERSION:		0.9

'**************************************************************************************************************
class Email

	'private members
	private p_recipients, p_sendersEmail, p_component, p_errorMsg, p_allTo
	private p_subject, p_body, p_sendersName, p_mailServerUsername, p_mailServerPassword, p_html
	
	'public members
	public onlyValidEmails	''[bool] indicates if the email addresses (sender & recipient) should be checked for correct syntax? default = true
	public mailServer		''[string] the host which is responsible for delivery. e.g. IP, hostname
	public mailer			''[object] gets the actual mailer object which has been wrapped by Email. Useful if some specific settings have to be made before sending.
	public dispatch			''[bool] indicates if the email should be dispatched when send() is called? default = true. useful to turn off if you only want to simulate sending on your dev env (to keep your inbox empty).
	
	public propertY get allTo ''[string] gets the email address to which all emails should be sent. useful on the dev env if you want to pipe all emails to one email address independently of the real recipient(s). this prevents that users would recieve an unwanted email which was send during the development. can be set using the AJAXED_EMAIL_ALLTO config
		allTo = p_allTo
	end property
	
	public property get subject ''[string] gets the subject
		subject = p_subject
	end property
	
	public property let subject(val) ''[string] sets the subject
		p_subject = val
		if component <> "" then
			mailer.subject = p_subject
		else
			exit property
		end if
	end property
	
	public property get body ''[string] gets the body
		body = p_body
	end property
	
	public property let body(val) ''[string] sets the body. if HTML is true then the body should be HTML markup
		p_body = val
		if component = "jmail.message" then
			if html then
				mailer.htmlBody = p_body
			else
				mailer.body = p_body
			end if
		elseif component = "persits.mailsender" then
			mailer.body = p_body
		elseif component = "cdo.message" then
			if html then
				mailer.htmlBody = p_body
			else
				mailer.textBody = p_body
			end if
		elseif component = "" then
			exit property
		else
			notSupported("body")
		end if
	end property
	
	public property get html ''[bool] indicates if the emails body will be interpreted as html
		html = p_html
	end property
	
	public property let html(val) ''[bool] sets the value indicating if the email body should be interpreted as html
		p_html = val
		if component = "persists.mailsender" then
			mailer.isHtml = p_html
		else
			exit property
		end if
	end property
	
	public property get sendersEmail ''[string] gets the senders email
		sendersEmail = p_sendersEmail
	end property
	
	public property let sendersEmail(val) ''[string] sets the senders email. if only valid emails allowed and its not a valid email then exception is thrown
		if onlyValidEmails and not str.isValidEmail(val) then lib.throwError("senders email is not a valid email address. you can turn the check off with onlyValidEmails")
		p_sendersEmail = val
		if component <> "" then
			mailer.from = p_sendersEmail
		else
			exit property
		end if
	end property
	
	public property get sendersName ''[string] gets the senders name
		sendersName = p_sendersName
	end property
	
	public property let sendersName(val) ''[string] sets the senders name.
		p_sendersName = val
		if component = "jmail.message" or component = "persits.mailsender" then
			mailer.fromName = p_sendersName
		elseif component = "cdo.message" then
			mailer.from = emailName(val, sendersEmail)
		elseif component = "" then
			exit property
		else
			notSupported("sendersName")
		end if
	end property
	
	public property get errorMsg ''[string] holds the errormessage if send() failed
		errorMsg = p_errorMsg
	end property
	
	public property get recipients ''[dictionary] gets already added recipients. key = auto ID, item = array(recipientstype, email, name)
		set recipients = p_recipients
	end property
	
	public property get supportedComponents ''[array] gets the supported mail components
		'Note: should all be lower-case!
		supportedComponents = array("jmail.message", "persits.mailsender", "cdo.message")
	end property
	
	public property get componentsOrdered ''[array] gets the loading order of the components (only necessary if more email components installed)
		componentsOrdered = lib.init(AJAXED_EMAIL_COMPONENTS, supportedComponents)
	end property
	
	public property get component ''[string] gets the name of the component which will be used for sending (which is supported). empty if there was no supported component found
		component = p_component
		if isEmpty(component) then
			p_component = lCase(lib.detectComponent(componentsOrdered))
			if isEmpty(p_component) then p_component = ""
		end if
		component = p_component
	end property
	
	public property get mailServerUsername ''[string] gets the username for the SMTP server.
		mailServerUsername = p_mailServerUsername
	end property
	
	public property let mailServerUsername(val) ''[string] sets the username for the SMTP server
		p_mailServerUsername = val
		if component = "jmail.message" then
			mailer.mailServerUserName = val
		elseif component = "persits.mailsender" then
			mailer.username = val
		elseif component = "cdo.message" then
			with mailer.configuration.fields
				.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = lib.iif(isEmpty(val), 0, 1)
				.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = val
				.update()
			end with
		elseif component = "" then
			exit property
		else
			notSupported("mailServername")
		end if
	end property
	
	public property get mailServerPassword ''[string] gets the password for the SMTP server.
		mailServerPassword = p_mailServerPassword
	end property
	
	public property let mailServerPassword(val) ''[string] sets the password for the SMTP server
		p_mailServerPassword = val
		if component = "jmail.message" then
			mailer.mailServerPassword = val
		elseif component = "persits.mailsender" then
			mailer.password = val
		elseif component = "cdo.message" then
			with mailer.configuration.fields
				.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = lib.iif(isEmpty(val), 0, 1)
				.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = val
				.update()
			end with
		elseif component = "" then
			exit property
		else
			notSupported("mailServerPassword")
		end if
	end property
	
	'******************************************************************************************************************
	'* constructor 
	'******************************************************************************************************************
	public sub class_Initialize()
		p_component = empty
		set mailer = nothing
		if component <> "" then set mailer = server.createObject(component)
		
		dispatch = lib.init(AJAXED_EMAIL_DISPATCH, true)
		p_allTo = lib.init(AJAXED_EMAIL_ALLTO, empty)
		if not isEmpty(allTo) then addRecipientToComponent "bcc", allTo, empty
		
		sendersEmail = lib.init(AJAXED_EMAIL_SENDER, empty)
		sendersName = lib.init(AJAXED_EMAIL_SENDER_NAME, empty)
		mailServer = lib.init(AJAXED_MAILSERVER, empty)
		mailServerUsername = lib.init(AJAXED_MAILSERVER_USER, empty)
		mailServerPassword = lib.init(AJAXED_MAILSERVER_PWD, empty)
		html = lib.init(AJAXED_HTML_EMAILS, false)
		
		onlyValidEmails = true
		set p_recipients = lib.newDict(empty)
		p_errorMsg = empty
	end sub 
	
	'******************************************************************************************************************
	'* destructor 
	'******************************************************************************************************************
	private sub class_terminate()
		set mailer = Nothing
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	STATIC! creates a new email instance with the content of a given TextTemplate
	'' @DESCRIPTION:	first line of the template is treated as the emails subject and the rest as the body.
	''					Dont forget to load the TextTemplate class for this.
	'' @PARAM:			template [TextTemplate]: template which is used within the email
	'' @RETURN:			[Email] a ready-to-use email instance where subject and body is set
	'******************************************************************************************************************
	public function newWith(template)
		if template is nothing then lib.throwError("No template givne.")
		set newWith = new Email
		with newWith
			.subject = template.getFirstLine()
			.body = template.getAllButFirstLine()
		end with
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	STATIC! sends an email with a given TextTemplate to a given recipient.
	'' @DESCRIPTION:	- a shortcut to quickly send an email
	''					- sender is the default one
	'' @PARAM:			template [TextTemplate]: template which is used within the email
	'' @PARAM:			recipient [string], [array]: if string then treated as email otherwise first
	''					value email and second the name
	'' @RETURN:			[string] error message if could not send otherwise empty
	'******************************************************************************************************************
	public default function quickSend(template, recipient)
		quickSend = empty
		set m = (new Email).newWith(template)
		if not isArray(recipient) then recipient = array(recipient)
		if ubound(recipient) < 1 then lib.throwError("recipient must be either an array with email and name or a string with email")
		if not m.addRecipient("TO", recipient(0), recipient(1)) then lib.throwError("recipients email is not a valid email.")
		if not m.send() then quickSend = m.errorMsg
		set m = nothing
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	sends the email message
	'' @DESCRIPTION:	- if sending failed then check errorMsg property for a detailed error
	'' @RETURN:			[bool] true if could be sent otherwise false
	'******************************************************************************************************************
	public function send()
		send = false
		if p_recipients.count <= 0 then p_errorMsg = "no recipient(s) email given"
		if sendersEmail = "" then p_errorMsg = "no sender email given"
		if p_errorMsg = "" then
			send = true
			if component = "persists.mailsender" then
				if dispatch then send = sendWithAspEmail()
			elseif component = "jmail.message" then
				if dispatch then send = sendWithJmail()
			elseif component = "cdo.message" then
				if dispatch then send = sendWithCDO()
			elseif component = "" then
				send = false
				p_errorMsg = "No supported email component found."
			else
				notSupported("send()")
			end if
		end if
		
		'log the email
		if not lib.logger.logsOnLevel(1) then exit function
		rTO = "To: " : rCC = "Cc: " : rBCC = "Bcc: "
		for each r in p_recipients.items
			recipient = emailName(r(2), r(1)) & ", "
			select case r(0)
				case "to" : rTO = rTO & recipient
				case "cc" : rCC = rCC & recipient
				case "bcc" : rBCC = rBCC & recipient
			end select
		next
		logInfo = array( _
			"Email " & lib.iif(send, "sent.", "NOT sent (" & errorMsg & ").") & lib.iif(dispatch, "", " dispatching OFF."), _
			"From: " & emailName(sendersName, sendersEmail), _
			"Subject: " & subject, _
			rTO, rCC, rBCC, "", _
			body _
		)
		lib.logger.log 1, logInfo, "1;37"
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	returns a formatted string for a given name and its email address 
	'' @DESCRIPTION:	useful for displaying sender or recipient nicely. 
	'' @PARAM:			aName [string]: name to be used
	'' @PARAM:			anEmail [string]: email to be used
	'' @RETURN:			[string] format: "name" <email@host.com>
	'******************************************************************************************************************
	public function emailName(byVal aName, byVal anEmail)
		if aName = "" then aName = anEmail
		emailName = """" & aName & """ <" & anEmail & ">"
	end function
	
	'******************************************************************************************************************
	'* sendWithAspEmail 
	'******************************************************************************************************************
	private function sendWithAspEmail()
		mailer.charset = "utf-8"
		on error resume next
			sendWithAspEmail = mailer.send()
			if err <> 0 then
				sendWithAspEmail = false
				errDesc = err.description
			end if
		on error goto 0
		p_errorMsg = errDesc
	end function
	
	'******************************************************************************************************************
	'* sendWithCDO 
	'******************************************************************************************************************
	private function sendWithCDO()
		if mailServer <> "" then
			with mailer.configuration.fields
				.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
				.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = mailServer
				.update()
			end with
		end if
		on error resume next
			sendWithCDO = mailer.send() = 0
			if err <> 0 then
				sendWithCDO = false
				errDesc = err.description
			end if
		on error goto 0
		p_errorMsg = errDesc
	end function
	
	'******************************************************************************************************************
	'* sendWithJmail 
	'******************************************************************************************************************
	private function sendWithJmail()
		sendWithJmail = false
		if mailServer = "" then
			p_errorMsg = "no mailserver configured. set mailServer property directly or globally in the config."
			exit function
		end if
		with mailer
			.charset = "utf-8"
			.logging = true
			.silent = true
			'otherwise iso-8859-1 error in subject
			.ISOEncodeHeaders = false
		end with
		sendWithJmail = mailer.send(mailServer)
		if not sendWithJmail then p_errorMsg = mailer.errormessage
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Adds a recipient. Can be used multiple times
	'' @DESCRIPTION:	- recipients property holds all added recipients
	'' @PARAM:			recipientsType [string]: Define the type of to. TO, CC, BCC
	'' @PARAM:			email [string]: Recipients email. Required! if onlyValidEmails is true then
	''					the recipient is only added if the email is a syntactically valid one
	'' @PARAM:			name [string]: Recipients name. if empty then email is used as name
	'' @RETURN:			[bool] true if the recipient has been added. false if not.
	''					Its not added if only valid emails are required and the email was not valid.
	'******************************************************************************************************************
	public function addRecipient(recipientsType, byVal email, byVal name)
		addRecipient = true
		recipientsType = lCase(recipientsType)
		if not lib.contains(array("to", "cc", "bcc"), recipientsType) then lib.throwError("Wrong recipientstype. Only to, cc or bcc.")
		if onlyValidEmails then addRecipient = str.isValidEmail(email)
		if not addRecipient then exit function
		
		if name = "" then name = email
		p_recipients.add lib.getUniqueID(), array(recipientsType, email, name)
		if isEmpty(allTo) then addRecipientToComponent recipientsType, email, name
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	helper which adds more than one recipient seperated by ";". Only email values. Name will be the email
	'' @PARAM:			recipientType [string]: type of the recipient. TO, CC or BCC
	'' @PARAM:			emails [string]: recipients emails seperated by ";"
	'**********************************************************************************************************
	public sub addRecipients(recipientType, emails)
		arr = split(emails, ";")
		for i = 0 to ubound(arr)
			addRecipient recipientType, arr(i), arr(i)
		next
	end sub
	
	'**********************************************************************************************************
	'* addRecipientToComponent 
	'**********************************************************************************************************
	private function addRecipientToComponent(recipientsType, email, name)
		if component = "persits.mailsender" then
			if recipientsType = "to" then
				mailer.addAddress email, name
			elseif recipientsType = "cc" then
				mailer.addCC email, name
			elseif recipientsType = "bcc" then
				mailer.addBCC email, name
			end if
		elseif component = "jmail.message" then
			if recipientsType = "to" then
				mailer.addRecipient email, name
			elseif recipientsType = "cc" then
				mailer.addRecipientCC email, name
			elseif recipientsType = "bcc" then
				mailer.addRecipientBCC email, name
			end if
		elseif component = "cdo.message" then
			if recipientsType = "to" then
				mailer.to = mailer.to & emailName(name, email) & "; "
			elseif recipientsType = "cc" then
				mailer.cc = mailer.cc & emailName(name, email) & "; "
			elseif recipientsType = "bcc" then
				mailer.bcc = mailer.bcc & emailName(name, email) & "; "
			end if
		elseif component = "" then
			exit function
		else
			notSupported("addRecipient")
		end if
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Adds an attachment. Can be used mutliple times
	'' @PARAM:			filename [string]: the name and path of your file on the server (absolute). e.g. c:\www\images\logo.gif
	'' @PARAM:			inline [bool]: if true the attachment will be added as an inline attachment
	'' @PARAM:			contentType [string]: attachments content type. leave empty if you don't want to specify it explicitly
	'' @RETURN:			[int] A unique ID that can be used to identify this attachment. This is useful if you are 
	''					embedding images in the email, body. Then you need to refer to it with <img src="cid:xxxx">
	'******************************************************************************************************************
	public function addAttachment(fileName, inline, contentType)
		addAttachment = lib.getUniqueID()
		if component = "persits.mailsender" then
			if contentType <> empty then notSupported("contentType of addAttachment()")
			if inline then
				mailer.addEmbeddedImage fileName, addAttachment
			else
				mailer.addAttachment(fileName)
			end if
		elseif component = "jmail.message" then
			if contentType = empty then
				addAttachment = mailer.addAttachment(fileName, inline)
			else
				addAttachment = mailer.addAttachment(fileName, inline, contentType)
			end if
		elseif component = "cdo.message" then
			if inline then notSupported("addAttachment(inline)")
			mailer.addAttachment(filename)
		elseif component = "" then
			exit function
		else
			notSupported("addAttachment")
		end if
	end function
	
	'******************************************************************************************************************
	'* notSupported 
	'******************************************************************************************************************
	private function notSupported(member)
		lib.throwError("'" & member & "' is not supported for mail component '" & component & "'")
	end function

end class
%>
