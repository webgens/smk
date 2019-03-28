<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'FacebookOAuth.asp - FacebookOAuth 콜백 페이지
'Date		: 2019.01.08
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





DIM MemberNum
DIM UserID
DIM DelFlag
DIM SNSChangeFlag
DIM MemberFlag
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'



Dim ResponseText
Dim HTTP_Object

Set HTTP_Object = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
With HTTP_Object
		'API 통신 Timeout 을 30초로 지정
		.SetTimeouts 30000, 30000, 30000, 30000
		.Open "POST", "https://graph.facebook.com/v2.12/oauth/access_token", False
		.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		.Send "client_id="& FACEBOOK_LOGIN_CLIENTID &"&redirect_uri="& HOME_DOMAIN_HTTS &"/API/FacebookOAuth.asp&client_secret="& FACEBOOK_LOGIN_APPSECRET &"&code="&request("code")
		.WaitForResponse


		If .Status = 200 Then
			ResponseText = .ResponseText
		Else
			ResponseText = ""
		End If
End With

Dim access_token
Dim token_type
Dim expires_in

Dim Read_Data
If ResponseText <> "" Then
		Set Read_Data = New aspJSON
		Read_Data.loadJSON(ResponseText)
		With Read_Data
				access_token	= .data("access_token")
				token_type		= .data("token_type")
				expires_in		= .data("expires_in")
		End With
End If

Dim resultCode
DIM ID
DIM Email
DIM KName


If access_token <> "" Then

		With HTTP_Object
				'API 통신 Timeout 을 30초로 지정
				.SetTimeouts 30000, 30000, 30000, 30000
				.Open "GET", "https://graph.facebook.com/me", False
				.SetRequestHeader "Content-Type", "application/json; charset=UTF-8"
				.SetRequestHeader "Authorization", "Bearer " & access_token
				.Send "access_token="&access_token&"&fields=id,name,email"
				.WaitForResponse

				If .Status = 200 Then
						ResponseText = .ResponseText
				Else
						ResponseText = ""
				End If
		End With


		If ResponseText <> "" Then
				Set Read_Data = New aspJSON
				Read_Data.loadJSON(ResponseText)
				With Read_Data
						ID		= .data("id")
						KName	= .data("name")
						IF .data("email") <> "" THEN
								Email	 = .data("email")
						ELSE
								Email	 = ""
						END IF
						'# Response.Write .data("id") & "/<br>"
						'# Response.Write .data("name") & "/<br>"
						'# Response.Write .data("email") & "/<br>"
						'# response.End
				End With
		End If

	




		SET oConn		 = ConnectionOpen()							'# 커넥션 생성
		SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


		'# SNS ID 체크
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Member_SNS_Select_By_SNSID"
	
				.Parameters.Append .CreateParameter("@SNSKind",	 adChar,	 adParamInput,  1, "F")
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
						Response.Cookies("SNS_Kind")		= Encrypt("F")


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
				Response.Cookies("SNS_Kind")		= Encrypt("F")
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
			var val = "<%=U_NUM%>///<%=ID%>///<%=Email%>///<%=Kname%>///F";
			APP_HistoryBack_SNS_Login(val);
		</script>
<%
ELSE
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=페이스북 로그인 오류 입니다.<br />다시 시도해 주십시오.&Script=APP_PopupHistoryBack();"
		Response.End
END IF
%>