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
            <h1 class="hidden">SHOEMARKER</h1>
            <button type="button" class="btn-goback" onclick="history.back()">
				<span class="hidden">이전 화면으로 돌아가기</span>
			</button>
            <span class="tit">SHOEMARKER</span>
            <button type="button" class="btn-basket" onclick="APP_GoUrl('/ASP/Order/CartList.asp')">
				<span class="hidden">장바구니</span>
				<span class="some" id="GNB_CartCount">0</span>
			</button>
            <button type="button" class="btn-srch">
				<span class="hidden">통합검색</span>
			</button>
        </div>
    </header>