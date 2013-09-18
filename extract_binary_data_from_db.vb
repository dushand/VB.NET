Option Explicit

Const C_DPath = "D:\temp"
Const C_IMG_NAME = "img01.jpg"

Const C_SRV = "server01"
Const C_DTB = "db01"
Const C_USR = "user01"
Const C_PWD = "pwd01"

Const C_TBL = "tbl01"
Const C_FLD_ID = "id01"
Const C_FLD_IMG = "img01"

Dim oConn, sAppIDD, sSQL, oRst, oFldImg, f, oFSO ' , oFile, n

Set oConn = CreateObject("ADODB.Connection")
oConn.ConnectionString = "Provider=SQLOLEDB;Data Source=" & C_SRV & ";Initial Database=" & C_DTB & ";User ID=" & C_USR & ";Password=" & C_PWD & ";"
oConn.Open

For Each sAppIDD In Array("1234567", "1234568", "1234569")
	sSQL = "SELECT " & C_FLD_IMG & " FROM " & C_TBL & " WHERE " & C_FLD_ID & " = " & sAppIDD
	Set oRst = oConn.Execute(sSQL)

	While Not oRst.EOF
		oFldImg = oRst.Fields(0).Value
		If Not IsNull(oFldImg) Then
			SaveFile
		End If
		oRst.MoveNext
	Wend
	oRst.Close
	Set oRst = Nothing
Next
oConn.close
Set oConn = Nothing


Sub SaveFile()
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	If Not oFSO.FolderExists(C_DPath & "\" & sAppIDD) Then
		oFSO.CreateFolder C_DPath & "\" & sAppIDD
		oFSO.CreateFolder C_DPath & "\" & sAppIDD & "\supp_doc"
	End If
	' save file without external object, slower
	'Set oFile = oFSO.CreateTextFile(C_IMG_NAME, True)
	'For n = 1 to LenB(oFldImg)
	'	oFile.Write Chr(AscB(MidB(oLOB, n, 1)))
	'Next
	'oFile.Close
	'Set oFile = Nothing
	SaveBinaryData C_DPath & "\" & sAppIDD & "\supp_doc\" & C_IMG_NAME, oFldImg
	Set oFSO = Nothing
End Sub

Sub SaveBinaryData(ByVal FileName As String, ByVal ByteArray As Object)
    Const adTypeBinary = 1
    Const adSaveCreateOverWrite = 2

    'Create Stream object
    Dim BinaryStream
    BinaryStream = CreateObject("ADODB.Stream")

    'Specify stream type - we want To save binary data.
    BinaryStream.Type = adTypeBinary

    'Open the stream And write binary data To the object
    BinaryStream.Open()
    BinaryStream.Write(ByteArray)

    'Save binary data To disk
    BinaryStream.SaveToFile(FileName, adSaveCreateOverWrite)

    BinaryStream.Close()
    BinaryStream = Nothing
End Sub

' faster file save, from http://www.motobit.com/tips/detpg_read-write-binary-files/
Function SaveBinaryData(FileName, ByteArray)
	Const adTypeBinary = 1
	Const adSaveCreateOverWrite = 2

	'Create Stream object
	Dim BinaryStream
	Set BinaryStream = CreateObject("ADODB.Stream")

	'Specify stream type - we want To save binary data.
	BinaryStream.Type = adTypeBinary

	'Open the stream And write binary data To the object
	BinaryStream.Open
	BinaryStream.Write ByteArray

	'Save binary data To disk
	BinaryStream.SaveToFile FileName, adSaveCreateOverWrite

	BinaryStream.Close
	Set BinaryStream = Nothing
End Function


