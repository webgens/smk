<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Ipin.asp - IPIN 본인인증 시작
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
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM i

DIM SMode
DIM JoinType
DIM Name
DIM UserID

DIM sSiteCode
DIM sSitePw
DIM sReturnURL
DIM sCPRequest
DIM IPIN_DLL
DIM clsIPINDll
DIM iRtn
DIM sEncReqData
DIM sRtnMsg

DIM UserName
DIM DormancyFlag
DIM NewAgreementFlag
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


SMode			 = sqlFilter(Request("SMode"))			'# 인증목적 : MemberJoin(회원가입) / SearchID(아이디찾기) / SearchPwd(비밀번호찾기) / DormancyRelease(휴면계정해제)
Name			 = sqlFilter(Request("Name"))
UserID			 = sqlFilter(Request("UserID"))


IF SMode = "" THEN
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=본인 인증 목적 값이 없습니다.<br />다시 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
		Response.End
'# 회원가입일 경우
ELSEIF SMode = "MemberJoin" THEN
		JoinType		 = TRIM(Decrypt(Request.Cookies("JOIN_TYPE")))
'# 비밀번호 찾기일 경우
ELSEIF SMode = "SearchPwd" THEN
		Response.Cookies("SW_NAME")		 = Encrypt(Name)
		Response.Cookies("SW_USERID")	 = Encrypt(UserID)
'# SNS 간편로그인 정회원 전환
ELSEIF SMode = "JoinChgMem" THEN
		JoinType		 = TRIM(Decrypt(Request.Cookies("JOIN_TYPE")))
END IF


IF SMode = "DormancyRelease" THEN
		UserName		 = TRIM(Decrypt(Request.Cookies("TEMP_UNAME")))
		DormancyFlag	 = TRIM(Decrypt(Request.Cookies("TEMP_DOR")))
		NewAgreementFlag = TRIM(Decrypt(Request.Cookies("TEMP_NEW")))


		IF DormancyFlag = "N" THEN
				IF NewAgreementFlag = "N" THEN
						Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=휴면계정 해제 처리 되었습니다.<br />신규 약관에 동의하여 주십시오.&Script=APP_PopupHistoryBack_Move('/ASP/Member/NewAgreement.asp');"
						Response.End
				ELSE
						IF U_ID <> "" THEN
								Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=휴면계정 해제 처리 되었습니다.&Script=APP_PopupHistoryBack_Move('/');"
								Response.End
						ELSE
								Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=" & UserName & "님은 휴면계정 해제 처리 되었습니다.&Script=APP_PopupHistoryBack_Move('/');"
								Response.End
						END IF
				END IF
		END IF
END IF



sReturnURL	= HOME_URL & "/Common/AuthIpin/IpinProcess.asp"
sCPRequest	= ""

SET clsIPINDll	= Server.CreateObject("IPINClient.Kisinfo")
	
sSiteCode		= IPIN_H_ID				'NICE로부터 부여받은 사이트 코드
sSitePw			= IPIN_H_PWD			'NICE로부터 부여받은 사이트 패스워드

clsIPINDll.fnRequestSEQ(sSiteCode)
sCPRequest = clsIPINDll.bstrRandomRequestSEQ

Response.Cookies("CPREQUEST") = sCPRequest

iRtn = clsIPINDll.fnRequest(sSiteCode, sSitePw, sCPRequest, sReturnURL)

IF (iRtn = 0) THEN
	
		'fnRequest 함수 처리시 업체정보를 암호화한 데이터를 추출합니다.
		'추출된 암호화된 데이타는 당사 팝업 요청시, 함께 보내주셔야 합니다.
		sEncReqData = clsIPINDll.bstrRequestCipherData
		
		sRtnMsg = "정상 처리되었습니다."
	
ELSEIF (iRtn = -9) THEN
		sRtnMsg = "입력값 오류 : fnRequest 함수 처리시, 필요한 4개의 파라미터값의 정보를 정확하게 입력해 주시기 바랍니다."
		sEncReqData = ""

		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=본인인증에 실패하였습니다.<br />본인인증 실패 사유 : 요청정보 암호화 오류[-9]&Script=APP_PopupHistoryBack();"
		Response.End
ELSE
		sRtnMsg = "iRtn 값 확인 후, NICE신용평가정보 개발 담당자에게 문의해 주세요."
		sEncReqData = ""

		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=본인인증에 실패하였습니다.<br />본인인증 실패 사유 : 요청정보 암호화 오류[" & iRtn & "]&Script=APP_PopupHistoryBack();"
		Response.End
END IF
    
SET clsIPINDll = NOTHING
%>
<html>
<head>
	<title>NICE신용평가정보 - IPIN 본인인증</title>
</head>
<body>
	<!-- 본인인증 서비스 팝업을 호출하기 위해서는 다음과 같은 form이 필요합니다. -->
	<form name="form_chk" method="post" action="https://cert.vno.co.kr/ipin.cb">
		<input type="hidden" name="m" value="pubmain">						<!-- 필수 데이타로, 누락하시면 안됩니다. -->
		<input type="hidden" name="enc_data" value="<%= sEncReqData %>">		<!-- 위에서 업체정보를 암호화 한 데이타입니다. -->
	    
	    <!-- 업체에서 응답받기 원하는 데이타를 설정하기 위해 사용할 수 있으며, 인증결과 응답시 해당 값을 그대로 송신합니다.
	    	 해당 파라미터는 추가하실 수 없습니다. -->
		<input type="hidden" name="param_r1" value="<%=SMode %>">
		<input type="hidden" name="param_r2" value="<%=JoinType %>">
		<input type="hidden" name="param_r3" value="">
	    
	</form>
	<script type="text/javascript">
		document.form_chk.submit();
	</script>
</body>
</html>