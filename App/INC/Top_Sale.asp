</head>

<body>
    
	<!-- Skip Nav -->
    <a href="#container" id="skipnav" class="skipnav">본문 바로가기</a>

    <!-- Header -->
    <header id="header" class="header">
        <div class="headline">
            <h1 class="primary-logo">
                <a href="#">SHOEMARKER</a>
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
        <section class="wrap-sort">
            <div class="sort" id="OrderType" onclick="OrderByLayerOpen();" style="width:50%;">
                <button type="button" id="orderbytype"><%=SortText%></button>
            </div>
            <div class="search" id="Search" onclick="SmartSearchLayerOpen();" style="width:50%;">
                <button type="button">스마트검색</button>
            </div>
        </section>
    </nav>
