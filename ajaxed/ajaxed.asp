<!--#include virtual="/ajaxedConfig/config.asp"-->
<!--#include file="class_localization/localization.asp"-->
<!--#include file="class_library/library.asp"-->
<!--#include file="class_dataContainer/dataContainer.asp"-->
<!--#include file="class_stringOperations/stringOperations.asp"-->
<!--#include file="class_stringBuilder/stringBuilder.asp"-->
<!--#include file="class_JSON/JSON.asp"-->
<!--#include file="class_database/database.asp"-->
<!--#include file="class_logger/logger.asp"-->
<!--#include file="class_ajaxedPage/ajaxedPage.asp"-->
<%
set lib = new Library
set lib.logger = new Logger
set local = new Localization
set str = new StringOperations
set db = new Database
%>