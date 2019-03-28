<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<script>
	location.href='https://kauth.kakao.com/oauth/authorize?client_id=<%=KAKAO_LOGIN_CLIENTID%>&redirect_uri=<%=HOME_DOMAIN_HTTS%>/API/KakaoOAuth.asp&response_type=code&state=<%=Server.UrlEncode("kakaologin")%>';
</script>