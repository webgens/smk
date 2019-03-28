    <div id="moveTop" class="move-top">
        <button type="button"><span class="hidden">top</span></button>
    </div>
    <article id="bnb" class="bnb">
        <ul class="grid">
            <li class="block archived"><a href="javascript:void(0)" onclick="location.href='/';" class="home">HOME</a></li>
            <li class="block"><a href="javascript:void(0)" onclick="GetCategory1();" class="menu">전체메뉴</a></li>
            <li class="block"><a href="javascript:void(0)" onclick="location.href='/ASP/Mypage/';" class="personal">마이 페이지</a></li>
            <li class="block"><a href="javascript:APP_GoUrl('/ASP/Mypage/MyPickList.asp');" class="favorite">FAVORITE</a></li>
<%
If U_Num = "" Then
	wQuery = " WHERE D.GuestInfo = '" & U_GuestInfo & "' "
Else
	wQuery = " WHERE D.MemberNum = " & U_Num & " OR D.GuestInfo = '" & U_GuestInfo & "' "
End If
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Latest_Select_By_wQuery_For_Top30"
		.Parameters.Append .CreateParameter("@wQuery", adVarChar, adParamInput, 1000, wQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

Dim Footer_Product_Latest_Count
Footer_Product_Latest_Count = oRs.RecordCount

If oRs.EOF Then
%>

            <!-- 최근 본 상품이 없을 때 -->
            <li class="block"><a href="javascript:common_msgPopOpen('', '최근 본 상품이 없습니다.');" class="recently">최근이 없을 때</a></li>
            <!-- 최근 본 상품이 없을 때 -->
<%
Else	
%>
            <!-- *** 수정 *** 190118 : 최근 본 상품  무한 롤링-->
            <!-- 최근 본 상품이 있을 때 -->
            <li class="block relatedItem">
                <div class="swiper-container">
                    <div class="swiper-wrapper">
					<% 
					Do While Not oRs.EOF	
					%>
                        <div class="swiper-slide">
                            <a href="javascript:open_ProductLatest();" class="recentlyHas"><img src="<%=oRs("ImageUrl")%>" alt="<%=oRs("ProductName")%>"></a>
                            <!-- 썸네일 사이즈 24px * 24px-->
                        </div>
					<%
						oRs.MoveNext
					Loop	
					%>
                    </div>
                </div>

            </li>
            <!-- 최근 본 상품이 있을 때 -->
<%
End If
oRs.Close	
%>
            <!-- *** 수정 *** 190118 : 최근 본 상품  무한 롤링-->
        </ul>
    </article>
	<footer id="footer" class="footer">
		<section class="notification" id="notification">
			<h1 class="subject">공지사항</h1>
			<p class="text" id="FooterNoticeList"></p>
		</section>
		<ul class="quick-link">
			<li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Customer/index.asp')">고객센터</a></li>
			<li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Customer/PartnerShip.asp')">입점/제휴 문의</a></li>
			<li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Customer/Store.asp')">매장찾기</a></li>
			<li><a href="javascript:void(0)" onclick="APP_PopupGoUrl('/ASP/Member/GradeInfo.asp', '0')">등급혜택</a></li>
		</ul>
		<button type="button" class="btn-more">슈마커몰 정보확인</button>
		<section class="company-detail">
			<div class="inner">
				<address>㈜에스엠케이티앤아이 (06210) 서울특별시 강남구 테헤란로 306, 카이트타워 7층</address>
				<p>
					<span>대표이사 : 안영환</span>
					<span>사업자등록번호 : 105-86-14706</span>
					<span>통신판매업등록번호 : 2009-서울강남-00623</span>
					<span>고객센터 : <a href="tel:080-030-2809">080-030-2809</a></span>
					<span>FAX : 02-711-9042 / 02-719-8956</span>
				</p>
				<ul class="grid">
					<li class="block"><a href="javascript:openExternal('http://www.ftc.go.kr/bizCommPop.do?wrkr_no=1058614706&apv_perm_no=');">사업자정보 확인</a></li>
					<li class="block"><a href="javascript:PolicyView(18);">이용약관</a></li>
					<li class="block"><a href="javascript:PolicyView(19);">개인정보처리방침</a></li>
				</ul>
				<p class="copyright">ⓒ Shoemarker All Rights Reserved.</p>
			</div>
		</section>
	</footer>



 	<!-- 약관 팝업 -->
	<div class="area-dim" id="PolicyPopup" style="display:none;z-index:120;">
	</div>
	<!-- 약관 팝업 -->

	<script type="text/javascript">
		var relatedItem = new Swiper('.relatedItem .swiper-container', {
			slidesPerView: 1,
			<% If Footer_Product_Latest_Count > 1 Then %>loop: true,<% End If %>
			autoplay: {
				delay: 2000,
				disableOnInteraction: false,
			},
			observer: true,
			observeParents: true
		});
	</script>