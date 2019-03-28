<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'AddressList.asp - 마이페이지 > 회원정보 > 배송지 관리
'Date		: 2018.12.17
'Update		: 
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'

'# 페이지 코드----------------------------------------------------------------------------'
DIM PageCode1, PageCode2, PageCode3, PageCode4
PageCode1 = "05"
PageCode2 = "05"
PageCode3 = "02"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/SubCheckID.asp" -->

<%

'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn						'# ADODB Connection 개체
DIM oRs							'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

DIM i
DIM j
DIM x
DIM y

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절


'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
	<style type="text/css">
		#OrderMenu .selector { margin-bottom: 0; }
		#OrderMenu .selector.is-focus .btn-list:after { background: url("/images/ico/ico_arrow_u2.png")no-repeat; background-size: 100% auto; }
	</style>
    <script type="text/javascript" src="/JS/dev/mypage.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
	<script type="text/javascript" src="/JS/dev/login.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
<%TopSubMenuTitle = "회원정보"%>
<!-- #include virtual="/INC/TopSub.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="sub_content">

            <div class="wrap-mypage">
				<div style="height:8px"></div>


				
                <div id="OrderMenu" class="ly-title accordion">
                    <div class="selector">
	                    <button type="button" class="btn-list clickEvt" data-target="OrderMenu">배송지관리</button>
					</div>
					<div class="option my-recode">
						<ul>
							<li><a href="/ASP/Mypage/MyMemberShip.asp">나의 멤버십</a></li>
							<li><a href="/ASP/Mypage/AddressList.asp">배송지관리</a></li>
							<li><a href="javascript:common_PopOpen('DimDepth1','MyInfoModify');">나의 정보 수정</a></li>
							<li><a href="/ASP/Mypage/SnsList.asp">SNS 계정설정</a></li>
						</ul>
					</div>
                </div>



                <div class="mypage-membership">

                    <section id="contentList_2" class="accord-mypage">
                        <div class="ly-content1" id="getMyAddrList">



							<form name="MyAddrListForm" id="MyAddrListForm">
							<input type="hidden" name="addrType" />
							<input type="hidden" name="idx" />
                            <div class="h-line">
                                <h2 class="h-level4">배송지 목록</h2>
                                <span class="h-date is-right">
                                    <button type="button" class="button-ty3 ty-bd-black" onclick="insert_MyAddr('insert','')">
                                        <span class="icon ico-add">배송지 추가하기</span>
                                </button>
                                </span>
                            </div>
<%
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_MyAddress_Select_By_MemberNum"

		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParaminput, , U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing


IF oRs.EOF THEN
%>
							<div class="deliver-list">
                                <p class="non_tit">등록된 정보가 없습니다.</p>
                            </div>
<%
ELSE
	Do While Not oRs.EOF
		IF oRs("MainFlag") = "Y" THEN
%>
							<div class="deliver-list">
                                <p class="tit"><%=oRs("AddressName") %></p>
								<div class="mypage">
                                    <span class="badge">기본</span>
                                </div>
                                <div class="address">
                                    <p class="">[<%=oRs("ReceiveZipCode") %>]</p>
                                    <p class=""><%=oRs("ReceiveAddr1") %> <%=oRs("ReceiveAddr2") %></p>
                                </div>
                                <div class="info-wrap">
                                    <span class="holder">받는분 : <%=oRs("ReceiveName") %></span>
                                    <span class="tel">연락처 : <%=oRs("ReceiveHP") %></span>
                                </div>
                                <button type="button" class="button-ty2 is-expand ty-bd-gray" onclick="insert_MyAddr('modify','<%=oRs("idx")%>')">수정</button>
                            </div>
<%
			ELSE
%>
							<div class="deliver-list">
                                <p class="tit"><%=oRs("AddressName") %></p>
                                <div class="address">
                                    <p class="">[<%=oRs("ReceiveZipCode") %>]</p>
                                    <p class=""><%=oRs("ReceiveAddr1") %> <%=oRs("ReceiveAddr2") %></p>
                                </div>
                                <div class="info-wrap">
                                    <span class="holder">받는분 : <%=oRs("ReceiveName") %></span>
                                    <span class="tel">연락처 : <%=oRs("ReceiveHP") %></span>
                                </div>
								<div class="buttongroup is-space">
                                    <button type="button" class="button-ty2 is-expand ty-bd-gray" onclick="insert_MyAddr('modify','<%=oRs("idx")%>')">수정</button>
                                    <button type="button" class="button-ty2 is-expand ty-bd-gray" onclick="chg_MainFlag('<%=oRs("idx") %>');">기본 배송지로 지정</button>
                                </div>
                                <div class="right-circle">
                                    <button type="button" class="closebtn" onclick="del_MyAddr('<%=oRs("idx") %>');">
                                        <span class="hidden">삭제</span>
                                    </button>
                                </div>
                            </div>
<%
			END IF
		oRs.MoveNext
	Loop
END IF
oRs.close
%>
							</form>


                        </div>
                    </section>


                </div>
            </div>
        </div>
    </main>



<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->


<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
