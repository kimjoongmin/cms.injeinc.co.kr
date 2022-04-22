<%@ CODEPAGE=65001%>
<%response.charset = "utf-8"%>
<!--#include virtual = "/include/db_open.asp"-->
<!--#include virtual = "/include/function.asp"-->
<%	
	Name		= FilterWord(request("Name"))
	Tel			= FilterWord(request("Tel"))
	Email		= FilterWord(request("Email"))
	Memo		= Write_txt(FilterWord(request("Memo")))
	auto_num	= FilterWord(request("auto_num"))

	If Name = "" Or Tel = "" Or Email = "" Or auto_num = "" Then
		ftnjsAlertMsgUrl "잘못된 경로로 접근하셨습니다. 다시 확인부탁드립니다.", "/"
		response.end
	End If
	
	chk_num		= request.cookies("chk_num")
		
	if abs(int(chk_num) - int(auto_num)) > 0  Then
		ftnjsAlertMsgUrl "자동등록방지를 위한 숫자가 일치하지 않습니다.", "/"
		response.End
	end If

	'## 변수처리 ##
	IP				= page_info(1)
	MaxNumber		= SR_FUN_getMaxFromDB ("tb_cms", "sn", "")
	state			= "N"

	'-----------쿼리 작성--------------------------------------------------------------
	SQL = "insert into tb_cms(sn,Name,Tel,Email,Memo,state,ip,regist_day) values ('"
	SQL = SQL &  MaxNumber        	& "','"

	SQL = SQL &  Name          	& "','"
	SQL = SQL &  Tel         	& "','"
	SQL = SQL &  Email          & "','"
	SQL = SQL &  Memo			& "','"

	SQL = SQL &  state			& "','"
	SQL = SQL &  ip				& "','"
	SQL = SQL &  now()			& "')"

	DB.execute(SQL)

	email = "pocas4356@injeinc.co.kr" '관리자 이메일 등록
	'email = "a3613@injeinc.co.kr" '관리자 이메일 등록

	'Call SendMail("관리자 <recruit@injeinc.co.kr>", email, "인재INC", "CMS 구매문의가 들어왔습니다.","")

	
	ftnjsAlertMsgUrl "구매문의가 신청되었습니다. 확인 후 연락드리겠습니다.","/"
%>
<!--#include virtual="/include/db_closs.asp"-->


