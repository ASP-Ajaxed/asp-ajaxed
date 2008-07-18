<!--#include file="../class_testFixture/testFixture.asp"-->
<%
set tf = new TestFixture
tf.debug = true
tf.run()

sub test_1()
	lib.logger.clearLogs()
	lib.logger.logLevel = 1
	tf.assertNot lib.fso.fileExists(server.mappath(lib.logger.logFile)), "log file should be gone after clearLogs()"
	tf.assert not lib.logger is nothing, "lib.logger not available"
	
	lib.logger.log 1, "some log msg", 31
	tf.assertInFile lib.logger.logFile, "[31msome log msg", "logger.log() didnt work properly (should contain ascii styles)"
	
	lib.logger.debug("testing the debug logging")
	tf.assertInFile lib.logger.logFile, "testing the debug logging", "logger.debug() didnt work"
	
	lib.logger.warn("some warning")
	tf.assertInFile lib.logger.logFile, "some warning", "logger.warn() didnt work"
	
	lib.logger.info("some info")
	tf.assertInFile lib.logger.logFile, "some info", "logger.info() didnt work"
	
	lib.logger.error("some error")
	tf.assertInFile lib.logger.logFile, "some error", "logger.error() didnt work"
end sub

sub test_2()
	lib.logger.logLevel = 0
	lib.logger.debug("notInIt1")
	lib.logger.warn("notInIt1")
	lib.logger.info("notInIt1")
	lib.logger.error("notInIt1")
	tf.assertNotInFile lib.logger.logFile, "notInIt1", "logile should not contain any logs if logLevel is 0 (disabled)"
	lib.logger.clearLogs()
	
	lib.logger.logLevel = 1
	lib.logger.debug("debugMSG")
	lib.logger.warn("warnMSG")
	lib.logger.info("infoMSG")
	lib.logger.error("errorMSG")
	tf.assertInFile lib.logger.logFile, "debugMSG", "logile should contain 'debug' msg if level is set to debug (1)"
	tf.assertInFile lib.logger.logFile, "warnMSG", "logile should contain 'warn' msg if level is set to debug (1)"
	tf.assertInFile lib.logger.logFile, "infoMSG", "logile should contain 'info' msg if level is set to debug (1)"
	tf.assertInFile lib.logger.logFile, "errorMSG", "logile should contain 'error' msg if level is set to debug (1)"
	lib.logger.clearLogs()
	
	lib.logger.logLevel = 2
	lib.logger.debug("debugMSG")
	lib.logger.warn("warnMSG")
	lib.logger.info("infoMSG")
	lib.logger.error("errorMSG")
	tf.assertNotInFile lib.logger.logFile, "debugMSG", "logile should NOT contain 'debug' msg if level is set to info (2)"
	tf.assertInFile lib.logger.logFile, "warnMSG", "logile should contain 'warn' msg if level is set to info (2)"
	tf.assertInFile lib.logger.logFile, "infoMSG", "logile should contain 'info' msg if level is set to info (2)"
	tf.assertInFile lib.logger.logFile, "errorMSG", "logile should contain 'error' msg if level is set to info (2)"
	lib.logger.clearLogs()
	
	lib.logger.logLevel = 4
	lib.logger.debug("debugMSG")
	lib.logger.warn("warnMSG")
	lib.logger.info("infoMSG")
	lib.logger.error("errorMSG")
	tf.assertNotInFile lib.logger.logFile, "debugMSG", "logile should NOT contain 'debug' msg if level is set to warn (4)"
	tf.assertNotInFile lib.logger.logFile, "infoMSG", "logile should NOT contain 'info' msg if level is set to warn (4)"
	tf.assertInFile lib.logger.logFile, "warnMSG", "logile should contain 'warn' msg if level is set to warn (4)"
	tf.assertInFile lib.logger.logFile, "errorMSG", "logile should contain 'error' msg if level is set to warn (4)"
	lib.logger.clearLogs()
	
	lib.logger.logLevel = 8
	lib.logger.debug("debugMSG")
	lib.logger.warn("warnMSG")
	lib.logger.info("infoMSG")
	lib.logger.error("errorMSG")
	tf.assertNotInFile lib.logger.logFile, "debugMSG", "logile should NOT contain 'debug' msg if level is set to error (8)"
	tf.assertNotInFile lib.logger.logFile, "infoMSG", "logile should NOT contain 'info' msg if level is set to error (8)"
	tf.assertNotInFile lib.logger.logFile, "warnMSG", "logile should NOT contain 'warn' msg if level is set to error (8)"
	tf.assertInFile lib.logger.logFile, "errorMSG", "logile should contain 'error' msg if level is set to error (8)"
	lib.logger.clearLogs()
end sub

sub test_3()
	'log utf8 chars
	lib.logger.logLevel = 1
	lib.logger.debug chrw(352)
end sub
%>