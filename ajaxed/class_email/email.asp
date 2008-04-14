<%
'**************************************************************************************************************

'' @CLASSTITLE:		Email
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2008-04-14 17:26
'' @CDESCRIPTION:	Represents an email. Create an instance and use send() to deliver the message
''					- use createNew() to create an email instance using a template
'' @REQUIRES:		-
'' @VERSION:		0.1

'**************************************************************************************************************
class Email

	private mailer, p_errorMessage, component
	
	'public members
	public subject			''[string] Email Subject
	public fromEmail		''[string] Senders Email. default = default email from config
	public fromName			''[string] Senders Name. default = default name from config
	public body				''[string] Body of the email. If htmlEmail true then it should be html-code
	public html				''[bool] Should the body be HTML encoded? default = false
	
	public property get errorMessage() ''[string] holds a detailed error message if the send() failed
		errormessage = p_errorMessage
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		component = lib.detectComponent(array("jmail.message", "persits.mailsender"))
		if isEmpty(component) then lib.throwError("Could not find any supported email component on the server.")
		set mailer = server.createObject(loadedComponent)
		fromEmail = lib.init(AJAXED_EMAIL_SENDER, empty)
		fromName = lib.init(AJAXED_EMAIL_SENDER_NAME, empty)
		smtpUsername = lib.init(AJAXED_EMAIL_USERNAME, empty)
		smtpPassword = lib.init(AJAXED_EMAIL_PASSWORD, empty)
		html = false
		body = empty
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Adds a recipient to the email-object. Use this method as often as you want to add recipients
	''					To your email.
	'' @DESCRIPTION:	if the name is empty then email is used as name
	'' @PARAM:			email [string]: recipients email
	'' @PARAM:			name [string]: recipients name
	'' @PARAM:			toWhat [string]: Define the type of to. E.g. CC, BCC, TO, etc.
	'******************************************************************************************************************
	public sub addRecipient(toWhat, email, name)
		if name = "" then name = email
		if component = "persits.mailsender" then
			select case ucase(toWhat)
				case "TO"
					mailer.addAddress email, name
				case "CC"
					mailer.addCC email, name
				case "BCC"
					mailer.addBCC email, name
			end select
		elseif component = "jmail.message" then
			select case ucase(toWhat)
				case "TO"
					mailer.addRecipient email, name
				case "CC"
					mailer.addRecipientCC email, name
				case "BCC"
					mailer.addRecipientBCC email, name
			end select
		end if
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Sends the email
	'' @RETURN:			[bool] true if successfully send otherwise false. if could not sent then a detailed errormessage
	''					can be found in the errormessage property.
	'******************************************************************************************************************
	public function send()
		send = false
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	Creates a new email with a given template
	'' @DESCRIPTION:	the first line of the template is used a subject and the rest as body
	'' @PARAM:			template [TextTemplate]: the template for the email
	'' @RETURN			[Email] a new instance of email with the subject and the body set
	'******************************************************************************************************************
	public function createNew(template)
		if lcase(typename(template)) <> "texttemplate"  then lib.throwError("a TextTemplate needs to be provided")
		set createNew = new Email
		createNew.subject = template.getFirstLine()
		createNew.body = template.getAllButFirstLine()
	end function

end class
%>