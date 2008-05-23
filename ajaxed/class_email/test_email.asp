﻿<% AJAXED_EMAIL_COMPONENTS = empty %>
<!--#include file="../class_testFixture/testFixture.asp"-->
<!--#include file="email.asp"-->
<%
set tf = new TestFixture
tf.debug = true
tf.run()

sub setup()
	if (new Email).allTo = "" then lib.throwError("Email.allTo property needs to be configured. piping all test emails to one person!")
end sub

'tries to send an email with each installed component
'if no components are installed then nothing will be tested ;)
sub test_1()
	components = (new Email).supportedComponents
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
			tf.assert e.send(), "email could not be sent with detected component '" & e.component & "' " & "(" & e.errorMsg & ")"
		end if
	next
end sub

'tries to send an emil with a component 
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