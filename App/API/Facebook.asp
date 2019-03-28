<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- 페이스북 로그인 -->
<script type="text/javascript">
	location.href='https://www.facebook.com/v2.12/dialog/oauth?client_id=<%=FACEBOOK_LOGIN_CLIENTID%>&redirect_uri=<%=HOME_DOMAIN_HTTS%>/API/FacebookOAuth.asp&state=<%=Server.UrlEncode("facebooklogin")%>&scope=public_profile,email';
</script>
<!-- 페이스북 로그인 -->
