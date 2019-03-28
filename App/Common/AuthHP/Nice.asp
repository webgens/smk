<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Nice.asp - 나이스 본인인증 시작
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
<!-- #include virtual="/Common/ProgID1.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM i

DIM SMode
DIM JoinType
DIM Name
DIM UserID

DIM clsCPClient
DIM iRtn
DIM sEncData
DIM sPlainData
DIM sAuthType
DIM sRequestNO
DIM sSiteCode
DIM sSitePassword
DIM sReturnUrl
DIM sErrorUrl
DIM popgubun
DIM sGender

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



SET clsCPClient  = SERVER.CREATEOBJECT("CPClient.Kisinfo")
	
sSiteCode		 = NICE_H_ID				'NICE로부터 부여받은 사이트 코드
sSitePassword    = NICE_H_PWD				'NICE로부터 부여받은 사이트 패스워드

sAuthType		 = "M"						'없으면 기본 선택화면, X: 공인인증서, M: 핸드폰, C: 카드
popgubun		 = "N"						'Y : 취소버튼 있음, N : 취소버튼 없음
sGender			 = ""						'없으면 기본 선택 값, 0 : 여자, 1 : 남자 

	    
'CheckPlus(본인인증) 처리 후, 결과 데이타를 리턴 받기위해 다음예제와 같이 http부터 입력합니다.
sReturnUrl		 = HOME_URL & "/Common/AuthHP/NiceSuccess.asp"			'성공시 이동될 URL
sErrorUrl		 = HOME_URL & "/Common/AuthHP/NiceFail.asp"				'실패시 이동될 URL

sRequestNO		 = "REQ0000000001"										'요청 번호, 이는 성공/실패후에 같은 값으로 되돌려주게 되므로
																		'업체에 적절하게 변경하여 쓰거나, 아래와 같이 생성한다.
iRtn = clsCPClient.fnRequestNO(sSiteCode)

IF iRtn = 0 THEN
		sRequestNO = clsCPClient.bstrRandomRequestNO
		Response.Cookies("REQ_SEQ") = sRequestNO		'해킹등의 방지를 위하여 세션을 쓴다면, 세션에 요청번호를 넣는다.
END IF

sPlainData		 = fnGenPlainData(sRequestNO, sSiteCode, sAuthType, sReturnUrl, sErrorUrl, popgubun, sGender)
	
'실제적인 암호화
iRtn			 = clsCPClient.fnEncode(sSiteCode, sSitePassword, sPlainData)

IF iRtn = 0 THEN
		sEncData = clsCPClient.bstrCipherData
ELSE
		'# RESPONSE.WRITE "요청정보_암호화_오류:" & iRtn & "<br>"
		' -1 : 암호화 시스템 에러입니다.
		' -2 : 암호화 처리오류입니다.
		' -3 : 암호화 데이터 오류입니다.
		' -4 : 입력 데이터 오류입니다.

		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=본인인증에 실패하였습니다.<br />본인인증 실패 사유 : 요청정보 암호화 오류.&Script=APP_PopupHistoryBack();"
		Response.End
END IF

SET clsCPClient = Nothing




'**************************************************************************************
'문자열 생성 
'**************************************************************************************  					          	
FUNCTION fnGenPlainData(aRequestNO, aSiteCode, aAuthType, aReturnUrl, aErrorUrl, popgubun, GENDER)
		DIM retPlainData
		'입력 파라미터로 plaindata 생성 			
		retPlainData  = "7:REQ_SEQ" & fnGetDataLength(aRequestNO) & ":" & aRequestNO & _
										"8:SITECODE" & fnGetDataLength(aSiteCode) & ":" & aSiteCode & _
										"9:AUTH_TYPE" & fnGetDataLength(aAuthType) & ":" & aAuthType & _
										"7:RTN_URL" & fnGetDataLength(aReturnUrl) & ":" & aReturnUrl & _
										"7:ERR_URL" & fnGetDataLength(aErrorUrl) & ":" & aErrorUrl	& _	
										"11:POPUP_GUBUN" & fnGetDataLength(popgubun) & ":" & popgubun & _
										"6:GENDER" & fnGetDataLength(sGender) & ":" & sGender
		fnGenPlainData = retPlainData		

END FUNCTION 

'**************************************************************************************
'입력파라미터의 문자열길이 추출	
'**************************************************************************************  					          	
FUNCTION fnGetDataLength(aData)		
		DIM iData_len, i
		IF (LEN(aData) > 0) THEN
				FOR i = 1 TO LEN(aData)
						IF (ASC(mid(aData,i,1)) < 0) THEN	'한글인경우
							iData_len = iData_len + 2
						ELSE								'한글이아닌경우
							iData_len = iData_len + 1
						END IF
				NEXT
		ELSE
				iData_len = 0
		END IF
			
		fnGetDataLength = iData_len
END FUNCTION
%>

<html>
<head>
	<title>NICE신용평가정보 - CheckPlus 본인인증</title>
</head>
<body>
	<!-- 본인인증 서비스 팝업을 호출하기 위해서는 다음과 같은 form이 필요합니다. -->
	<form name="form_chk" method="post" action="https://check.namecheck.co.kr/checkplus_new_model4/checkplus.cb">
		<input type="hidden" name="m" value="checkplusSerivce">						<!-- 필수 데이타로, 누락하시면 안됩니다. -->
		<input type="hidden" name="EncodeData" value="<%= sEncData %>">		<!-- 위에서 업체정보를 암호화 한 데이타입니다. -->
	    
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