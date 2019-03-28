	<style type="text/css">
		.gnb .gnb-list li>a:after{content: '';display: block;position: absolute;bottom: 0;left: 0;width: 100%;height: 3px;background-color: transparent;transition: width .1s ease-in-out }
		.gnb .gnb-list li.current>a:after{width: 100%;background-color: #ff201b;}
	</style>
</head>

<body>

	<style type="text/css">
		#loading {position:fixed;border-radius: 5px;display:inline-block;top:50%;left:50%;margin-top:-23px;margin-left:-23px;z-index:1000;background:#000000;display:none;text-align:center; opacity: 0.5; -ms-filter: alpha(opacity=50); filter: alpha(opacity=50);}
	</style>
	<div class="loading_dim"></div>
	<div id="loading">
		<img src="/Images/loading.gif" class="show-page-loading-msg" style="width:46px" alt="처리중" />
	</div>

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
	
	<!-- Navigation -->
	<nav id="gnb" class="gnb">
        <div class="wrap-gnb">
            <div class="area-gnb">
                <ul class="gnb-list">
					<li <% If PageCode1 = "00" Then %>class="current"<% End If %>><a href="/">홈</a></li>
					<li <% If PageCode1 = "BR" Then %>class="current"<% End If %>><a href="/ASP/Product/Brands.asp">BRANDS</a></li>
					<li <% If PageCode1 = "T1" Then %>class="current"<% End If %>><a href="/ASP/Product/Top100.asp">TOP100</a></li>
					<li <% If PageCode1 = "SL" Then %>class="current"<% End If %>><a href="/ASP/Product/Sale.asp">SALE</a></li>
					<li <% If PageCode1 = "TD" Then %>class="current"<% End If %>><a href="/ASP/Product/Today.asp">TODAY'S DEAL</a></li>
					<li <% If PageCode1 = "EV" Then %>class="current"<% End If%>><a href="/ASP/Event/EventList.asp">이벤트</a></li>
					<li><a href="/ASP/Street306/">STREET306</a></li>
					<li><a href="/ASP/ShoemarkerOnly/">ONLY</a></li>
                </ul>
            </div>
        </div>
	</nav>
