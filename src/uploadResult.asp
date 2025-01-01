<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!-- #include file="Upload.asp" -->
<html>
<head>
   <title> Temp Text </title>
   <meta name="viewport" content="width=device-width, initial-scale=1">
   <meta name="Author" content="Austin">
</head>
<body>
<%
' ��ȡ·��
dim ufp,filePath,extList
extList = "|ini|md|mp3|m4a|bmp|jpg|jpeg|png|zip|7z|rar|pdf|doc|docx|xls|xlsx|ppt|pptx|epub"
filePath= "./temp/"

const rowSplit = "|"
Function myFileName(fileName,count)
	myFileName = URLDecode(Split(fileName,"|")(count))
End Function
function URLDecode(strIn)
    URLDecode = ""
    Dim sl: sl = 1
    Dim tl: tl = 1
    Dim key: key = "%"
    Dim kl: kl = Len(key)

    sl = InStr(sl, strIn, key, 1)
    Do While sl>0
        If (tl=1 And sl<>1) Or tl<sl Then
            URLDecode = URLDecode & Mid(strIn, tl, sl-tl)
        End If

        Dim hh, hi, hl
        Dim a
        Select Case UCase(Mid(strIn, sl+kl, 1))
            Case "U":                  'Unicode URLEncode
            a = Mid(strIn, sl+kl+1, 4)
            URLDecode = URLDecode & ChrW("&H" & a)
            sl = sl + 6

            Case "E":                   'UTF-8 URLEncode
            hh = Mid(strIn, sl+kl, 2)
            a = Int("&H" & hh)          'ascii��
            If Abs(a)<128 Then
                sl = sl + 3
                URLDecode = URLDecode & Chr(a)
            Else
                hi = Mid(strIn, sl+3+kl, 2)
                hl = Mid(strIn, sl+6+kl, 2)
                a = ("&H" & hh And &H0F) * 2 ^12 Or ("&H" & hi And &H3F) * 2 ^ 6 Or ("&H" & hl And &H3F)
                If a<0 Then a = a + 65536
                URLDecode = URLDecode & ChrW(a)
                sl = sl + 9
            End If
        Case Else:                      'Asc URLEncode
            hh = Mid(strIn, sl+kl, 2)   '��λ
            a = Int("&H" & hh)          'ascii��
            If Abs(a)<128 Then
            sl = sl + 3
            Else
            hi = Mid(strIn, sl+3+kl, 2) '��λ
            a = Int("&H" & hh & hi)     '��ascii��
            sl = sl + 6
            End If
            URLDecode = URLDecode & Chr(a)
        End Select

        tl = sl
        sl = InStr(sl, strIn, key, 1)
    Loop

    URLDecode = URLDecode & Mid(strIn, tl)
End function
function myConvert(strIn)
	' ���� ADODB.Stream ����
	Set stream = Server.CreateObject("ADODB.Stream")

	' ����������Ϊ�ı�����
	stream.Type = 2 

	' �����ַ���Ϊ UTF-8
	stream.Charset = "gbk"
	' ����
	stream.Open
	' ���������ַ���д����
	stream.WriteText strIn

	' ����ת��Ϊ ANSI ����
	stream.Position = 0

	stream.Charset = "utf-8"
	strOut = stream.ReadText
	' �ر���
	stream.Close

	' �ͷ�������
	Set stream = Nothing
	' ������
	myConvert = strOut
end function



Server.ScriptTimeout = 900

set Upload = new DoteyUpload

Upload.Upload()

if Upload.ErrMsg <> "" then 
Response.Write(Upload.ErrMsg)
Response.End()
end if

if Upload.Files.Count > 0 then
Items = Upload.Files.Items
end if

fileArr = Array()
ReDim fileArr(Upload.Files.Count-1)

expireFile = Upload.Form("expireFile")
fileName = Upload.Form("fileName")

count = 0
tick = day(now)&hour(now)&minute(now)&second(now)
for each File in Upload.Files.Items
    ' currFileName = myConvert(File.FileName )
    currFileName = myFileName(fileName,count )

	upfilename = split(currFileName,".")
	upfileext = Lcase(upfilename(ubound(upfilename)))
 ' 
	if InStr(extList,"|"&upfileext) > 0  then
		ufp= tick & "-" & currFileName
		file.saveas Server.mappath(filePath&ufp)
		' �ļ���һֱ <size> ���
		fileArr(count) =  ufp & "<"& file.fileSize& ">"
		
		count = count + 1
	else
		Response.Write("�ļ���׺" & upfileext & "������Ҫ��")
		Response.End
	end if
	' �����¼�����ݿ�
	' Application("tempText")
	' file.fileSize&",'"&path&ufp&"','"&session("username")&"','"&upfileext
next
Application(tick) = Join(fileArr,rowSplit)
Set upload=nothing
Response.Write("�Ѿ��ɹ��ϴ�")

if ufp<>"" then %>
<script language="JavaScript">
window.location.href="../cb.asp?tick=<%=tick%>&expireFile=<%=expireFile%>"
</script>
<%
end if
%>

</body></html>