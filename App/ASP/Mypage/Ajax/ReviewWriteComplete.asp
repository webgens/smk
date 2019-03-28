<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'ReviewWriteComplete.asp - 상품후기 작성 완료 페이지
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

DIM ReviewType
DIM ReviewPoint		: ReviewPoint	= 0
DIM TotalPoint		: TotalPoint	= 0
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


ReviewType			= sqlFilter(Request("ReviewType"))




IF ReviewType = "P" THEN
		ReviewPoint	= MALL_REVIEW_POINT_P
ELSE
		ReviewPoint	= MALL_REVIEW_POINT_B
END IF




SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성



IF U_MFLAG = "Y" THEN
		Set oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection = oConn
				.CommandType = adCmdStoredProc
				.CommandText = "USP_Admin_EShop_Member_Select_By_MemberNum"

				.Parameters.Append .CreateParameter("@MemberNum",	adInteger,	adParamInput, ,		U_NUM)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		Set oCmd = Nothing

		IF NOT oRs.EOF THEN
				TotalPoint		= oRs("Point")
		END IF
		oRs.Close
END IF


Response.Write "OK|||||"
%>					
        <div class="area-dim" style="z-index:101"></div>

        <div class="area-pop">
            <div class="alert">
                <div class="tit-pop">
                    <p class="tit">상품후기 등록 완료</p>
                </div>

                <div class="container-pop">
                    <div class="contents">
                        <div class="ly-cont">
							<%IF U_MFLAG = "Y" AND CDbl(ReviewPoint) > 0 THEN%>
                            <p class="t-level4">
								상품 후기가 등록되어<br> 
								<em style="color:#e60917; font-size:14px;"><%=FormatNumber(ReviewPoint, 0)%><span>원</span></em>이 포인트로 적립되었습니다.<br />
								고객님의 총 보유 포인트가 <em style="color:#e60917; font-size:14px;"><%=FormatNumber(TotalPoint, 0)%>원</em>이 되었습니다. <br />
								(즉시 사용가능)
                            </p>
							<%ELSE%>
                            <p class="t-level4">상품 후기가 등록 되었습니다.</p>
							<%END IF%>
                        </div>
                    </div>
                    <div class="btns">
                        <button type="button" onclick="orderListReload()" class="button ty-red" style="width:100%">확인</button>
                    </div>
                </div>
            </div>
        </div>

<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>