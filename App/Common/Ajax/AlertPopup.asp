<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'AlertPopup.asp - 알림팝업 페이지
'Date		: 2018.11.16
'Update	: 
'/****************************************************************************************'

'//페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//---------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->

<%
'/****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM PopupID
DIM AlertType
DIM AlertTitle
DIM AlertMsg
DIM Button1Name
DIM Button2Name
DIM Button1Action
DIM Button2Action
DIM ButtonWidth
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


PopupID			= sqlFilter(Request("PopupID"))
AlertType		= sqlFilter(Request("AlertType"))
AlertTitle		= sqlFilter(Request("AlertTitle"))
AlertMsg		= sqlFilter(Request("AlertMsg"))
Button1Name		= sqlFilter(Request("Button1Name"))
Button2Name		= sqlFilter(Request("Button2Name"))
Button1Action	= sqlFilter(Request("Button1Action"))
Button2Action	= sqlFilter(Request("Button2Action"))
ButtonWidth		= sqlFilter(Request("ButtonWidth"))


IF AlertType = "AddCart" THEN
%>					
        <div class="area-dim" style="z-index:101"></div>

        <div class="area-pop">
            <div class="alert">
                <div class="tit-pop">
                    <p class="tit">장바구니 담기 완료</p>
                    <!--<button class="btn-hide-pop">닫기</button>-->
                </div>

                <div class="container-pop">
                    <div class="contents">
                        <div class="ly-cont">
                            <p class="t-level4">선택하신 상품을 장바구니에 담았습니다. <br> 장바구니로 이동하시겠습니까?</p>
                        </div>
                    </div>
                    <div class="btns">
                        <button type="button" onclick="closePop('msgPopup');APP_GoUrl('/ASP/Order/CartList.asp')" class="button ty-black">장바구니로 이동</button>
                        <button type="button" onclick="PageReload()" class="button ty-red">쇼핑 계속하기</button>
                    </div>
                </div>
            </div>
        </div>
<%
ELSE
%>
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="alert">
                <div class="tit-pop">
                    <p class="tit"><%=AlertTitle%></p>
                    <!--<button class="btn-hide-pop">닫기</button>-->
                </div>

                <div class="container-pop">
                    <div class="contents">
                        <div class="ly-cont">
                            <p class="t-level4"><%=AlertMsg%></p>
                        </div>
                    </div>
                    <div class="btns">
                        <button type="button" <%IF ButtonWidth <> "" THEN%>style="width:<%=ButtonWidth%>px"<%END IF%> onclick="<%=Button1Action%>; closePop('<%=PopupID%>')" class="button ty-black"><%=Button1Name%></button>
                        <button type="button" <%IF ButtonWidth <> "" THEN%>style="width:<%=ButtonWidth%>px"<%END IF%> onclick="<%=Button2Action%>; closePop('<%=PopupID%>')" class="button ty-red"><%=Button2Name%></button>
                    </div>
                </div>
            </div>
        </div>
<%
END IF
%>
