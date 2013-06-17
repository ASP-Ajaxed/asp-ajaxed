<%
'***************************************
' File:	         Upload.asp
' Author:        Jacob "Beezle" Gilley
' Email:         avis7@airmail.net
' Date:          12/07/2000
' Updated:       12/20/2002
' Modified by:   Will Bickford
' Email:         wbic16@hotmail.com
' Comments: The code for the Upload, CByteString, 
'    CWideString subroutines was originally 
'    written by Philippe Collignon...or so 
'    he claims. This script is provided
'    "AS-IS" without support of any kind.
' Taken from: http://www.asp101.com/articles/jacob/scriptupload.asp
'****************************************

Class FileUploader
	Public  Files
	Private mcolFormElem

	Public Property Get Form(sIndex)
		Form = ""
		If mcolFormElem.Exists(LCase(sIndex)) Then Form = mcolFormElem.Item(LCase(sIndex))
	End Property

	Public Default Sub Upload()
		Dim biData, sInputName
		Dim nPosBegin, nPosEnd, nPos, vDataBounds, nDataBoundPos
		Dim nPosFile, nPosBound

		biData = Request.BinaryRead(Request.TotalBytes)
		nPosBegin = 1
		nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))

		If (nPosEnd-nPosBegin) <= 0 Then Exit Sub

		vDataBounds = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
		nDataBoundPos = InstrB(1, biData, vDataBounds)

		Do Until nDataBoundPos = InstrB(biData, vDataBounds & CByteString("--"))

			nPos = InstrB(nDataBoundPos, biData, CByteString("Content-Disposition"))
			nPos = InstrB(nPos, biData, CByteString("name="))
			nPosBegin = nPos + 6
			nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(34)))
			sInputName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
			nPosFile = InstrB(nDataBoundPos, biData, CByteString("filename="))
			nPosBound = InstrB(nPosEnd, biData, vDataBounds)

			If nPosFile <> 0 And  nPosFile < nPosBound Then
				Dim oUploadFile, sFileName
				Set oUploadFile = New UploadedFile

				nPosBegin = nPosFile + 10
				nPosEnd =  InstrB(nPosBegin, biData, CByteString(Chr(34)))
				sFileName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
				oUploadFile.FileName = Right(sFileName, Len(sFileName)-InStrRev(sFileName, "\"))

				nPos = InstrB(nPosEnd, biData, CByteString("Content-Type:"))
				nPosBegin = nPos + 14
				nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
				oUploadFile.ContentType = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))

				nPosBegin = nPosEnd+4
				nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
				oUploadFile.FileData = MidB(biData, nPosBegin, nPosEnd-nPosBegin)

				If oUploadFile.FileSize > 0 Then Files.Add LCase(sInputName), oUploadFile
			Else
				nPos = InstrB(nPos, biData, CByteString(Chr(13)))
				nPosBegin = nPos + 4
				nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
				If Not mcolFormElem.Exists(LCase(sInputName)) Then mcolFormElem.Add LCase(sInputName), CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
			End If
			nDataBoundPos = InstrB(nDataBoundPos + LenB(vDataBounds), biData, vDataBounds)
		Loop
	End Sub

	Private Sub Class_Initialize()
		Set Files = Server.CreateObject("Scripting.Dictionary")
		Set mcolFormElem = Server.CreateObject("Scripting.Dictionary")
	End Sub

	Private Sub Class_Terminate()
		If IsObject(Files) Then
			Files.RemoveAll()
			Set Files = Nothing
		End If

		If IsObject(mcolFormElem) Then
			mcolFormElem.RemoveAll()
			Set mcolFormElem = Nothing
		End If
	End Sub

	'String to byte string conversion
	Private Function CByteString(sString)
		Dim nIndex
		For nIndex = 1 to Len(sString)
		   CByteString = CByteString & ChrB(AscB(Mid(sString,nIndex,1)))
		Next
	End Function

	'Byte string to string conversion
	Private Function CWideString(bsString)
		Dim nIndex
		CWideString =""
		For nIndex = 1 to LenB(bsString)
		   CWideString = CWideString & Chr(AscB(MidB(bsString,nIndex,1))) 
		Next
	End Function
End Class

Class UploadedFile
	Public ContentType
	Public FileName
	Public FileData
	
	Public Property Get FileSize()
		FileSize = LenB(FileData)
	End Property

	Public Sub SaveToDisk(sPath,sCFN)
		Dim oFS, oFile
		Dim nIndex

		If sPath = "" Or sCFN = "" Then Exit Sub
		'If FileData = "" Or FileName = "" Then Exit Sub
		If Mid(sPath, Len(sPath)) <> "\" Then sPath = sPath & "\"

		Set oFS = Server.CreateObject("Scripting.FileSystemObject")

		If Not oFS.FolderExists(sPath) Then Exit Sub

		Set oFile = oFS.CreateTextFile((sPath & sCFN), True, False)
		oFile.Write BufferContent(FileData)
		oFile.Close
	End Sub

	Public Sub SaveToDatabase(ByRef oField)
		If LenB(FileData) = 0 Then Exit Sub

		If IsObject(oField) Then
			oField.AppendChunk FileData
		End If
	End Sub
	
	'***********************************************************
	'Code for more efficient buffering
	' Original Code written by: Robbert Nix
	' Adapted and Modified by: Will Bickford
	' Date: 12/20/2002
	' Email: wbic16@hotmail.com
	' From: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=4&txtCodeId=7110
	
	Private Function BufferContent(data)
		Dim strContent(64)
		Dim i
	
		ClearString strContent
	
		For i = 1 To LenB(data)
			AddString strContent,Chr(AscB(MidB(data,i,1)))
		Next
	
		BufferContent = fnReadString(strContent)
	End Function
	
	Private Sub ClearString(part)
		Dim index
	
		For index = 0 to 64
			part(index)=""
		Next
	End Sub
	
	Private Sub AddString(part,newString)
		Dim tmp
		Dim index
	
		part(0) = part(0) & newString
	
		If Len(part(0)) > 64 Then
			index=0
			tmp=""
	
			Do
				tmp=part(index) & tmp
				part(index) = ""
				index = index + 1
			Loop until part(index) = ""
	
			part(index) = tmp
		End If
	End Sub
	
	Private Function fnReadString(part)
		Dim tmp
		Dim index
	
		tmp = ""
	
		For index = 0 to 64
			If part(index) <> "" Then
				tmp = part(index) & tmp
			End If
		Next
	
		FnReadString = tmp
	End Function
End Class
%>