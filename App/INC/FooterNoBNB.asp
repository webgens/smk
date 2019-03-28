    <div id="moveTop" class="move-top">
        <button type="button"><span class="hidden">top</span></button>
    </div>
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

		<ul class="quick-link">
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
