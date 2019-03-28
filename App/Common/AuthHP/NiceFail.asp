<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'NiceFail.asp - 나이스 본인인증 실패
'Date		: 2018.11.06
'Update		: 
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->

<%
DIM clsCPClient
DIM sEncodeData, sSiteCode, sSitePassword, sCipherTime, iRtn
DIM sRequestNumber             '요청 번호
DIM sErrorCode                 '인증 결과코드
DIM sAuthType                  '인증 수단
DIM sReserved1, sReserved2, sReserved3
DIM sResult

DIM sPlain
DIM iReturn
DIM sRequestNO


SET clsCPClient  = SERVER.CREATEOBJECT("CPClient.Kisinfo")

sEncodeData		 = Fn_checkXss(Request("EncodeData"), "encodeData")	

sSiteCode		= NICE_H_ID				'NICE로부터 부여받은 사이트 코드
sSitePassword   = NICE_H_PWD			'NICE로부터 부여받은 사이트 패스워드

iRtn = clsCPClient.fnDecode(sSiteCode, sSitePassword, sEncodeData)

IF iRtn = 0 THEN
		sPlain           = clsCPClient.bstrPlainData
		sCipherTime      = clsCPClient.bstrCipherDateTime

		iReturn			 = clsCPClient.fnGetAuthInfo("REQ_SEQ")
		sRequestNumber	 = clsCPClient.bstrAuthInfo
		
		iReturn			 = clsCPClient.fnGetAuthInfo("ERR_CODE")
		sErrorCode		 = clsCPClient.bstrAuthInfo
		
		iReturn			 = clsCPClient.fnGetAuthInfo("AUTH_TYPE")
		sAuthType		 = clsCPClient.bstrAuthInfo
		
		sRequestNO		 = sRequestNumber
ELSE
		sErrorCode = "요청정보_암호화_오류:" & iRtn
		' -1 : 암호화 시스템 에러입니다.
		' -4 : 입력 데이터 오류입니다.
		' -5 : 복호화 해쉬 오류입니다.
		' -6 : 복호화 데이터 오류입니다.
		' -9 : 입력 데이터 오류입니다.
		'-12 : 사이트 패스워드 오류입니다.
END IF
Set clsCPClient = Nothing

FUNCTION Fn_checkXss (CheckString, CheckGubun) 
		CheckString = trim(CheckString)
		CheckString = replace(CheckString,"<","&lt;")
		CheckString = replace(CheckString,">","&gt;")
		CheckString = replace(CheckString,"""","")  
		CheckString = replace(CheckString,"'","")   
		CheckString = replace(CheckString,"(","")
		CheckString = replace(CheckString,")","")
		CheckString = replace(CheckString,"#","")
		CheckString = replace(CheckString,"%","")
		CheckString = replace(CheckString,";","")
		CheckString = replace(CheckString,":","")
		CheckString = replace(CheckString,"-","")      
		CheckString = replace(CheckString,"`","")
		CheckString = replace(CheckString,"--","")
		CheckString = replace(CheckString,"\","")
		IF CheckGubun <> "encodeData" THEN	
				CheckString = replace(CheckString,"+","")
				CheckString = replace(CheckString,"=","")
				CheckString = replace(CheckString,"/","")
		END IF	
		Fn_checkXss = CheckString
END FUNCTION



Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=본인인증에 실패하였습니다.<br />" & sErrorCode & "&Script=APP_PopupHistoryBack();"
Response.End
%>
<!--
<html>
<head>
	<title>NICE신용평가정보 - CheckPlus 본인인증</title>
</head>
<body>
	<script type="text/javascript">
		alert("본인인증에 실패하였습니다.\n\n<%=sErrorCode%>");
		self.close();
	</script>
</body>
</html>
-->