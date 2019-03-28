</head>

<body>
	<!-- Skip Nav -->
	<a href="#container" id="skipnav" class="skipnav">본문 바로가기</a>
	
	<!-- Header -->
	<header id="header" class="header">
		<div class="headline">
		    <h1 class="primary-logo">
		        <a href="/">SHOEMARKER</a>
		    </h1>
		    <button type="button" class="btn-basket" onclick="APP_GoUrl('/ASP/Order/CartList.asp')">
				<span class="hidden">장바구니</span>
	            <span class="some" id="GNB_CartCount">0</span>
			</button>
		    <button type="button" class="btn-srch" onclick="TopSearch();">
				<span class="hidden">통합검색</span>
			</button>
		</div>
	</header>
	
	<nav id="gnb" class="gnb">
        <ul class="gnb-list">
			<%IF U_MFLAG = "Y" THEN%>
            <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/OrderList.asp')">쇼핑내역</a></li>
            <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/MyPickList.asp')"">MY슈마커</a></li>
            <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/CouponList.asp')"">쇼핑혜택</a></li>
            <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/MyMemberShip.asp')"">회원정보</a></li>
			<%ELSE%>
            <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/OrderList.asp')">쇼핑내역</a></li>
            <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/MyPickList.asp')">MY슈마커</a></li>
            <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/MemberShip.asp')">쇼핑혜택</a></li>
            <li><a href="javascript:void(0)" onclick="snsAlert();">회원정보</a></li>
			<%END IF%>
        </ul>
	</nav>
	<script type="text/javascript">
		function snsAlert() {
			openAlertLayer("alert", "정회원만 이용 가능합니다.", "closePop('alertPop', '');", "");
		}
	</script>
