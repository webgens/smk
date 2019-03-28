<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<script src="https://apis.google.com/js/api.js"></script>

<script>
	var url = 'https://accounts.google.com/o/oauth2/v2/auth?scope=https://www.googleapis.com/auth/userinfo.profile+https://www.googleapis.com/auth/userinfo.email&state=security_token%3D138r5719ru3e1%26url%3D<%=HOME_DOMAIN%>&redirect_uri=<%=HOME_DOMAIN%>/API/GoogleOAuth.asp&response_type=code&client_id=<%=GOOGLE_LOGIN_CLIENTID%>&prompt=consent&include_granted_scopes=true';
	location.href = url;
</script>
<%

'https://www.googleapis.com/auth/plus.login
'https://www.googleapis.com/auth/userinfo.email
%>