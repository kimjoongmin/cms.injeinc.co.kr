<%
	'===========================================================
' 함 수 명 : 필터링
' 설    명 :
' 입 력 값 :
' 사 용 예 : FilterWord (문자)
'===========================================================
Function FilterWord(SR_Word)
	SR_Word		= replace(SR_Word,"'","''")
	SR_Word		= replace(SR_Word,""""," ")

	'SR_Word		= ftnsafeStr(SR_Word)
	SR_Word		= ftnDBFilter(SR_Word)
	'SR_Word		= SR_SUB_FilterWord(SR_Word)
	'SR_Word		= ftnsafeStr(SR_Word)
	
	FilterWord = SR_Word
End Function

'===========================================================
' 함 수 명 : 필터링
' 설    명 : 공백제거 및 콤마를 | 이것으로 변환
' 입 력 값 :
' 사 용 예 : FilterWord (문자)
'===========================================================
Function FilterWord2(SR_Word)
	C_Word = Replace(C_Word, "select", "")
	C_Word = Replace(C_Word, "union", "")
	C_Word = Replace(C_Word, "insert", "")
	C_Word = Replace(C_Word, "update", "")
	C_Word = Replace(C_Word, "delete", "")
	C_Word = Replace(C_Word, "create", "")
	C_Word = Replace(C_Word, "alter", "")
	C_Word = Replace(C_Word, "drop", "")
	C_Word = Replace(C_Word, "exec", "")
	C_Word = Replace(C_Word, "set @", "")
	C_Word = Replace(C_Word, "1=1", "")
	C_Word = Replace(C_Word, "sysobject", "")
	C_Word = Replace(C_Word, "alert", "")
	C_Word = Replace(C_Word, "on error resume", "")
	C_Word = Replace(C_Word, "execute", "")
	C_Word = Replace(C_Word, "windows", "")
	C_Word = Replace(C_Word, "boot", "")
	C_Word = Replace(C_Word, "-1 or", "")
	C_Word = Replace(C_Word, "unexisting", "")
	C_Word = Replace(C_Word, "win.ini", "")
	C_Word = Replace(C_Word, "%20", "")

	FilterWord2 = SR_Word
End Function


Function FilterWord3(SR_Word)
	If SR_Word <> "" And SR_Word <> "undefined" then
		SR_Word		= replace(SR_Word,"'","''")
	End If

	'SR_Word		= ftnsafeStr(SR_Word)


	FilterWord3 = SR_Word
End Function

'===========================================================
' 함 수 명 : SQL injection 방어함수
' 설    명 :
' 입 력 값 :
' 사 용 예 : ftnDBFilter (문자)
'===========================================================
function ftnDBFilter(C_Word)

	C_Word = Replace(C_Word, "select", "")
	C_Word = Replace(C_Word, "union", "")
	C_Word = Replace(C_Word, "insert", "")
	C_Word = Replace(C_Word, "update", "")
	C_Word = Replace(C_Word, "delete", "")
	C_Word = Replace(C_Word, "create", "")
	C_Word = Replace(C_Word, "alter", "")
	C_Word = Replace(C_Word, "drop", "")
	C_Word = Replace(C_Word, "exec", "")
	C_Word = Replace(C_Word, "set @", "")
	C_Word = Replace(C_Word, "1=1", "")
	C_Word = Replace(C_Word, "sysobject", "")
	C_Word = Replace(C_Word, "alert", "")

	C_Word = Replace(C_Word, "on error resume", "")
	C_Word = Replace(C_Word, "execute", "")
	C_Word = Replace(C_Word, "windows", "")
	C_Word = Replace(C_Word, "boot", "")
	C_Word = Replace(C_Word, "-1 or", "")
	C_Word = Replace(C_Word, "--, #", "")
	C_Word = Replace(C_Word, "/* */", "")
	C_Word = Replace(C_Word, "or 1=1--", "")
	C_Word = Replace(C_Word, "-1' or", "")
	C_Word = Replace(C_Word, "../", "")
	C_Word = Replace(C_Word, "unexisting", "")
	C_Word = Replace(C_Word, "win.ini", "")

	C_Word = Replace(C_Word, "%20", "")
	C_Word = replace(C_Word," or ","")
	C_Word = replace(C_Word," Or ","")
	C_Word = replace(C_Word," oR ","")
	C_Word = replace(C_Word," OR ","")
	C_Word=  replace(C_Word," and ","")

	C_Word = Replace(C_Word, "--", "")
	C_Word = replace(C_Word,"^","")
	C_Word = replace(C_Word,">=","")
	C_Word = replace(C_Word,"<=","")
	'C_Word = replace(C_Word,">","")
	'C_Word = replace(C_Word,"<","")
	C_Word = replace(C_Word,"<>","")
	C_Word = replace(C_Word,";","")

	C_Word = Replace(C_Word, "dir", "")
	C_Word = Replace(C_Word, "config", "")
	C_Word = Replace(C_Word, "systeminfo", "")
	C_Word = Replace(C_Word, "shutdown", "") 

	ftnDBFilter = C_Word
end Function

function SR_SUB_FilterWord(C_Word)
	C_Word = replace(C_Word,"'","''")
	C_Word = replace(C_Word,"%20","")
	C_Word = replace(C_Word," or ","")
	C_Word = replace(C_Word," Or ","")
	C_Word = replace(C_Word," oR ","")
	C_Word = replace(C_Word," OR ","")
	C_Word=  replace(C_Word," and ","")
	C_Word=  replace(C_Word,"(","")
	C_Word=  replace(C_Word,")","")
	C_Word=  replace(C_Word,"\","")
	C_Word=  replace(C_Word,"/","")
	C_Word=  replace(C_Word,"+","")
	C_Word=  replace(C_Word,"|","")
    'C_Word=  replace(C_Word,",","")
	C_Word = replace(C_Word,"=","")
	C_Word = replace(C_Word,"--","")
	C_Word = replace(C_Word,">=","")
	C_Word = replace(C_Word,"<=","")
	C_Word = replace(C_Word,">","")
	C_Word = replace(C_Word,"<","")
	C_Word = replace(C_Word,"<>","")
	C_Word = replace(C_Word,";","")
	C_Word = replace(C_Word,":","")
	C_Word = replace(C_Word,"^","")
	C_Word = replace(C_Word,"""","")
	C_Word = replace(C_Word,"%","")
	C_Word = Replace(C_Word, "select", "")
	C_Word = Replace(C_Word, "union", "")
	C_Word = Replace(C_Word, "insert", "")
	C_Word = Replace(C_Word, "update", "")
	C_Word = Replace(C_Word, "delete", "")
	C_Word = Replace(C_Word, "create", "")
	C_Word = Replace(C_Word, "alter", "")
	C_Word = Replace(C_Word, "drop", "")
	C_Word = Replace(C_Word, "exec", "")
	C_Word = Replace(C_Word, "set @", "")
	C_Word = Replace(C_Word, "1=1", "")
	C_Word = Replace(C_Word, "sysobject", "")

	SR_SUB_FilterWord = C_Word

End Function


'===========================================================
' 함 수 명 : page_info(num)
' 설    명 : 페이지정보 보기
' 입 력 값 : 1->아이피주소,2->절대경로,3->현재페이지 꼬리말,4->이전페이지 전체주소
' 입 력 값 : 5->이전페이지 꼬리말, 6->도메인명
' 사 용 예 : page_info(1)
'===========================================================
Function page_info(num)

	 HTTP_REFERER	= Request.ServerVariables("HTTP_REFERER")
	 b				= instrrev(HTTP_REFERER,"?")

	 select case num
		case 1
			info = Request.ServerVariables("REMOTE_ADDR")
		case 2
			info = Request.ServerVariables("PATH_INFO")
		case 3
			info = Request.ServerVariables("QUERY_STRING")
		case 4
			info = HTTP_REFERER
		case 5
			info = mid(HTTP_REFERER,b+1)
		case 6
			info = Request.ServerVariables("HTTP_HOST")
	 end Select

	 page_info = info
end Function


Function Write_txt(str)
	If Not IsNull(str) Then

		'str		= ftnsafeStr(str)
		str		= replace(str, chr(13)&chr(10), "<BR>")

		write_txt = str
	End if
End Function

Function view_txt(str)
	If Not IsNull(str) Then

		str		= replace(str, chr(13)&chr(10), "<BR>")

		view_txt = str
	End if
End Function

'===========================================================
' 함 수 명 : SR_FUN_getMaxFromDB
' 설    명 : 테이블 가장큰값 리턴
' 입 력 값 : 테이블명,컬럼,조건문
' 사 용 예 : SR_FUN_getMaxFromDB (테이블명, "컬럼명", "조건문")
'===========================================================
Function SR_FUN_getMaxFromDB (TableName, Column, strsqlWhere)
	SQL = "SELECT Max(" & Column & ") FROM " & TableName &  strSqlWhere
'	Response.write
'	Response.end
	Set oRs = DB.Execute(SQL)

	IF oRS.BOF or oRS.EOF Or IsNULL(oRs(0)) Then
		MaxVal = 1
	Else
		MaxVal	= oRs(0)+1
	END IF
	oRs.close

	SR_FUN_getMaxFromDB = MaxVal
End Function


'===========================================================
' 함 수 명 : ftnjsAlertMsgUrl
' 설    명 : 알럿 메세지 띄우고 해당페이지로 이동
' 입 력 값 : 메세지
' 사 용 예 : ftnjsAlertMsgUrl "문구내용", "이동경로"
'===========================================================
Function ftnjsAlertMsgUrl(msg, url)
	response.write "<script language=javascript>" & vbCrLf
	response.write "<!--" & vbCrLf
	if msg <> "" then
		response.write "    window.alert('" & msg & "');" & vbCrLf
	end if
	if url <> "" then
		response.write "    location.href='" & url & "';" & vbCrLf
	end if
	response.write "// -->" & vbCrLf
	response.write "</script>"
end Function


Sub SendMail(QnA_EMail1,QnA_EMail2,QnA_Title,QnA_Content,sOption)
	Set objmail = Server.CreateObject("CDO.message")
	Set iConf   = objmail.Configuration

	With iConf.Fields
		 .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1
		 .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = "C:\Inetpub\mailroot\Pickup"
		 .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "127.0.0.1"
		 .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		 .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
		 .Update
	End With

	 objmail.From				 = QnA_EMail1
	 objmail.to					 = QnA_EMail2
	 objmail.Subject			 = QnA_Title
	 objmail.HTMLBody			 = QnA_Content
	 objmail.BodyPart.Charset     = "euc-kr"
	 objmail.HTMLBodyPart.Charset = "euc-kr"
	 objmail.Send

	 'Response.write "발송완료"
End Sub


%>