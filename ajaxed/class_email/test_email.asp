<% AJAXED_EMAIL_COMPONENTS = empty %>
<!--#include file="../class_testFixture/testFixture.asp"-->
<!--#include file="email.asp"-->
<%
allToExists = empty
set tf = new TestFixture
tf.debug = true
tf.run()

sub setup()
	if not isEmpty(allToExists) then exit sub
	allToExists = true
	if (new Email).allTo = "" then
		tf.info("Email.allTo must be set for the email tests. Thus email sending was not fully tested. set AJAXED_EMAIL_ALLTO in your config to fully test email sending.")
		allToExists = false
	end if
end sub

'tries to send an email with each installed component
'if no components are installed then nothing will be tested ;)
sub test_1()
	components = (new Email).supportedComponents
	sentEmails = 0
	for each c in components
		AJAXED_EMAIL_COMPONENTS = array(c)
		set e = new Email
		if e.component <> "" then
			e.subject = "Test"
			e.body = "test"
			e.sendersEmail = "tester@ajaxed.com"
			e.addRecipient "to", "some@someemail.com", "michal"
			e.addRecipient "cc", "someCC@someemail.com", "CCer"
			e.addRecipient "bcc", "someBCC@someemail.com", "BCCer"
			e.body = "test email with " & e.component
			e.addAttachment server.mappath("test_attachment.txt"), false, empty
			if not allToExists then e.dispatch = false
			sent = e.send()
			tf.assert sent, "email could not be sent with detected component '" & e.component & "' " & "(" & e.errorMsg & ")"
			if sent then sentEmails = sentEmails + 1
		end if
	next
	if sentEmails > 0 then
		if (new Email).dispatch then
			tf.info(sentEmails & " test emails should be in the inbox of '" & (new Email).allTo & "'")
		else
			tf.info(sentEmails & " test emails have been sent but not dispatched because dispatching is turned off. Turn dispatching on if full test of sending emails is required.'")
		end if
	end if
end sub

'tries to send an emil with a component which does not exist
sub test_2()
	AJAXED_EMAIL_COMPONENTS = array("no component")
	set e = new Email
	e.subject = "Test"
	e.sendersEmail = "test@tester.com"
	e.sendersName = "Test Tester"
	e.addRecipient "TO", "john@recipient.com", "John Recipient"
	e.body = "test"
	tf.assertNot e.send(), "it must be possible to send an email even if no component has been found"
	tf.assert e.errorMsg <> "", "after sending without component there must be an errormessage"
end sub
%>