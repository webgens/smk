<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderConfirm.asp - 주문 구매확정 폼 페이지
'Date		: 2019.01.03
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
IF U_NUM = "" AND N_NAME = "" THEN
		Response.Write "LOGIN|||||로그인 정보가 없습니다. 다시 로그인하여 주십시오."
		Response.End
END IF

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

DIM OrderCode
DIM Idx

DIM OrderStateNM
DIM ProductPoint
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderCode			= sqlFilter(Request("OrderCode"))
Idx					= sqlFilter(Request("Idx"))




IF OrderCode = "" OR Idx = "" THEN
		Response.Write "FAIL|||||선택한 주문상품이 없습니다."
		Response.End
END IF




SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성



'# 주문정보 체크
wQuery = ""
wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType = 'P' "
wQuery = wQuery & "AND A.OrderCode = '" & OrderCode & "' "
wQuery = wQuery & "AND A.Idx = " & Idx & " "
IF U_NUM <> "" THEN
		wQuery = wQuery & "AND B.UserID = '" & U_NUM & "' "
ELSE
		wQuery = wQuery & "AND B.OrderName = '" & N_NAME & "' AND B.OrderHp = '" & N_HP & "' AND B.OrderEmail = '" & N_EMAIL & "' "
END IF

sQuery = ""

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Order_Product_Select_For_Order_Detail"

		.Parameters.Append .CreateParameter("@WQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@SQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		OrderStateNM	= GetOrderState(oRs("OrderState"), oRs("CancelState1"), oRs("CancelState2"))

		IF InStr(OrderStateNM, "배송중") <= 0 AND InStr(OrderStateNM, "배송완료") <= 0 THEN
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||구매확정할 수 없는 주문상태 입니다."
				Response.End
		END IF

		ProductPoint	= oRs("ProductPoint")

ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||구매확정할 주문내역이 없습니다."
		Response.End
END IF
oRs.Close


Response.Write "OK|||||"
%>					
        <div class="area-dim" style="z-index:101"></div>

        <div class="area-pop">
            <div class="alert">
                <div class="tit-pop">
                    <p class="tit">구매확정</p>
                    <button type="button" onclick="closePop('DimDepth1')" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop">
                    <div class="contents">
                        <div class="ly-cont">
							<%IF U_MFLAG = "Y" AND CDbl(ProductPoint) > 0 THEN%>
                            <p class="t-level4">
								구매 확정 하시면<br> 
								<em style="color:#e60917; font-size:14px;"><%=FormatNumber(ProductPoint, 0)%><span>원</span></em>이 포인트로 적립됩니다.<br />
								구매 확정 하시겠습니까?
                            </p>
							<%ELSE%>
                            <p class="t-level4">구매 확정 하시겠습니까?</p>
							<%END IF%>
                        </div>
                        <div class="inf-type1" style="margin-top:25px; margin-bottom:0;">
                            <p class="tit">주의사항</p>
                            <ul>
                                <li class="bullet-ty1">구매확정 후의 상품에 대한 반품/교환문의는 고객센터를 통해서만 접수하실 수 있습니다.</li>
                            </ul>
                        </div>
                    </div>
                    <div class="btns">
                        <button type="button" onclick="closePop('DimDepth1')" class="button ty-black">취소</button>
                        <button type="button" onclick="orderConfirm('<%=OrderCode%>', '<%=Idx%>')" class="button ty-red">확인</button>
                    </div>
                </div>
            </div>
        </div>

<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>