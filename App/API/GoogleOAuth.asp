<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'GoogleOAuth.asp - 구글 콜백 페이지
'Date		: 2018.12.14
'Update		: 
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

<!-- #include virtual="/API/json_for_asp/aspJSON1.17.asp" -->

<!-- #include virtual="/INC/Header.asp" -->

<script language="javascript" runat="server" charset="utf-8">
	function replaceAll(strTemp, strValue1, strValue2){ 
		while (1) {
			if (strTemp.indexOf(strValue1) != -1)
				strTemp = strTemp.replace(strValue1, strValue2);
			else
				break;
		}
		return strTemp;
	}
	function unicodeToKor(a) {
		var str = a;
		return unescape(replaceAll(str, "\\", "%"));
	}
</script>
<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn						'# ADODB Connection 개체
DIM oRs							'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절
DIM sqlQuery


DIM ResponseText
DIM HTTP_Object

DIM AccessToken
DIM RefreshToken
DIM TokenType
DIM ExpiresIn
DIM state
DIM ErrorCode
DIM ErrorDesc
DIM ReadData

DIM ResultCode
DIM ID
DIM Email
DIM KName
DIM Gender
DIM BirthDay
DIM Age
DIM NickName
DIM ProfileImage

DIM MemberNum
DIM UserID
DIM DelFlag
DIM SNSChangeFlag
DIM MemberFlag
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'



SET HTTP_Object = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
WITH HTTP_Object
		'API 통신 Timeout 을 30초로 지정
		.SetTimeouts		 30000, 30000, 30000, 30000
		.Open				 "POST",			 "https://www.googleapis.com/oauth2/v4/token", False
		.SetRequestHeader	 "Content-Type",	 "application/x-www-form-urlencoded; charset=UTF-8"
		.Send				 "grant_type=authorization_code&client_id="& GOOGLE_LOGIN_CLIENTID &"&client_secret="& GOOGLE_LOGIN_CLIENTSECRET &"&redirect_uri="& HOME_DOMAIN &"/API/GoogleOAuth.asp&code=" & Request("code")
		.WaitForResponse

		IF .Status = 200 THEN
				ResponseText = .ResponseText
		ELSE
				ResponseText = ""
		END IF
END WITH

IF ResponseText <> "" THEN
		SET ReadData = New aspJSON
		ReadData.loadJSON(ResponseText)
		WITH ReadData
				AccessToken		 = .data("access_token")
				RefreshToken	 = .data("refresh_token")
				TokenType		 = .data("token_type")
				ExpiresIn		 = .data("expires_in")
				ErrorCode		 = .data("error")
				ErrorDesc		 = .data("error_description")
		END WITH
		SET ReadData = Nothing
END If


If AccessToken <> "" Then

		WITH HTTP_Object
				'API 통신 Timeout 을 30초로 지정
				.SetTimeouts 30000, 30000, 30000, 30000
				.Open				 "GET",						 "https://www.googleapis.com/oauth2/v2/userinfo", False
				.SetRequestHeader	 "Content-Type",			 "application/json; charset=UTF-8"
'				.SetRequestHeader	 "X-Naver-Client-Id",		 NAVER_LOGIN_CLIENTID
'				.SetRequestHeader	 "X-Naver-Client-Secret",	 NAVER_LOGIN_CLIENTSECRET
				.SetRequestHeader	 "Authorization",			 "Bearer " & AccessToken
				.Send
				.WaitForResponse

				IF .Status = 200 THEN
						ResponseText = .ResponseText
				ELSE
						ResponseText = ""
				END IF
		END With
		
		SET ReadData = New aspJSON
		ReadData.loadJSON(ResponseText)
		WITH ReadData
				ID		 = .data("id")
				Email	 = .data("email")
				KName	 = .data("name")
		END WITH
		SET ReadData = Nothing



		SET oConn		 = ConnectionOpen()							'# 커넥션 생성
		SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


		'# SNS ID 체크
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Member_SNS_Select_By_SNSID"
	
				.Parameters.Append .CreateParameter("@SNSKind",	 adChar,	 adParamInput,  1, "G")
				.Parameters.Append .CreateParameter("SNSID",	 adVarChar,	 adParamInput, 50, ID)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				MemberNum = oRs("MemberNum")
		ELSE
				'# SNS계정연결에 사용
				IF U_NUM <> "" THEN
						MemberNum = U_NUM
				ELSE
						oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing

						Response.Cookies("SNS_UID")		= Encrypt(ID)
						Response.Cookies("SNS_Email")	= Encrypt(Email)
						Response.Cookies("SNS_KName")	= Encrypt(KName)
						Response.Cookies("SNS_Kind")		= Encrypt("G")


						Response.Redirect "/ASP/Member/SnsGate.asp"
						Response.End
				END IF
		END IF
		oRs.Close



		'# 회원정보 체크
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Admin_EShop_Member_Select_By_MemberNum"
	
				.Parameters.Append .CreateParameter("@MemberNum", adInteger, adParamInput,  , MemberNum)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				UserID			 = oRs("UserID")
				DelFlag			 = oRs("DelFlag")
				SNSChangeFlag	 = oRs("SNSChangeFlag")
				MemberFlag		 = oRs("MemberFlag")

				Response.Cookies("SNS_UID")		= Encrypt(ID)
				Response.Cookies("SNS_Email")	= Encrypt(Email)
				Response.Cookies("SNS_KName")	= Encrypt(KName)
				Response.Cookies("SNS_Kind")		= Encrypt("G")
				Response.Cookies("SNS_UserID")	= Encrypt(UserID)
				Response.Cookies("SNS_UNUM")		= Encrypt(MemberNum)
		ELSE
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing


				Response.Redirect "/ASP/Member/SnsGate.asp"
				Response.End
		END IF
		oRs.Close


		IF DelFlag = "Y" THEN
				SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=회원님은 탈퇴하신 정보가 있습니다.<br />재가입하여 주시기 바랍니다.&Script=APP_PopupHistoryBack();"
				Response.End
		END IF





		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing
%>
		<script type="text/javascript">
			var val = "<%=U_NUM%>///<%=ID%>///<%=Email%>///<%=Kname%>///G";
			APP_HistoryBack_SNS_Login(val);
		</script>
<%
ELSE
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=구글 로그인 오류 입니다.<br />다시 시도해 주십시오.&Script=APP_PopupHistoryBack();"
		Response.End
END IF
%>