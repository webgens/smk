<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<script type="text/javascript" src="//static.nid.naver.com/js/naverLogin_implicit-1.0.3.js" charset="utf-8"></script>
<script type="text/javascript" src="//code.jquery.com/jquery-1.11.3.min.js"></script>
<script>
	var url = 'https://nid.naver.com/oauth2.0/authorize?client_id=<%=NAVER_LOGIN_CLIENTID%>&response_type=code&redirect_uri=<%=HOME_URL%>/API/NaverOAuth.asp&state=<%=Server.UrlEncode("naverlogin")%>';
	location.href = url;
</script>