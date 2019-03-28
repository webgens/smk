<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'ProductAdd.asp - 상품목록 추가 페이지
'Date		: 2018.12.27
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
DIM oRs1											'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM Seq
DIM SalePriceType
DIM ProductCode
DIM SizeCD
DIM OProductCode
DIM OSizeCD

DIM BrandName
DIM ProductName
DIM SalePrice
DIM EventProdNM		: EventProdNM	= ""
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


Seq				= sqlFilter(Request("Seq"))
SalePriceType	= sqlFilter(Request("SalePriceType"))
ProductCode		= sqlFilter(Request("ProductCode"))
SizeCD			= sqlFilter(Request("SizeCD"))
OProductCode	= sqlFilter(Request("OProductCode"))
OSizeCD			= sqlFilter(Request("OSizeCD"))

IF SalePriceType = "" THEN SalePriceType = "1"


IF Seq = "" OR ProductCode = "" OR SizeCD = "" THEN
		Response.Write "FAIL|||||선택한 상품이 없습니다."
		Response.End
END IF


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


'# 상품정보
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Select_By_ProductCode"

		.Parameters.Append .CreateParameter("@ProductCode", adInteger, adParamInput, , ProductCode)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		BrandName		= oRs("BrandName")
		ProductName		= oRs("ProductName")
		IF SalePriceType = "2" THEN
				SalePrice		= oRs("EmployeeSalePrice")
		ELSE
				SalePrice		= oRs("SalePrice")
		END IF
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||없는 상품 입니다."
		Response.End
END IF
oRs.Close


'# 사은품정보
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_SubProduct_Event_Select_By_ProductCode"

		.Parameters.Append .CreateParameter("@ProductCode",		 adInteger, adParaminput,		, ProductCode)
End WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		Do Until oRs.EOF
				IF EventProdNM = "" THEN
						EventProdNM		= oRs("EventProdNM")
				ELSE
						EventProdNM		= EventProdNM & ", " & oRs("EventProdNM")
				END IF

				IF oRs("Qty") > 1 THEN
						EventProdNM		= EventProdNM & "(" & oRs("Qty") & ")"
				END IF

				oRs.MoveNext
		Loop
END IF
oRs.Close


Response.Write "OK|||||"
%>					
                        <li class="Product<%=Seq%>">
							<input type="hidden" name="Seq"				value="<%=Seq%>" />
							<input type="hidden" name="ProductCode"		value="<%=ProductCode%>" />
							<input type="hidden" name="SizeCD"			value="<%=SizeCD%>" />
							<input type="hidden" name="ProductType"		value="P" />
							<input type="hidden" name="SalePriceType"	value="<%=SalePriceType%>" />
							<input type="hidden" name="SalePrice"		value="<%=SalePrice%>" />
                            <div class="selected-cont">
                                <div class="tit"><span class="brand-name"><%=BrandName%></span><span class="item-name"><%=ProductName%></span></div>
                                <div class="cont">
                                    <div class="amount">
									<%IF OProductCode = "" THEN%>
                                        <button type="button" onclick="changeQty('<%=Seq%>', -1)" class="btn-minus dis">-</button>
                                        <span class="product-length">1</span>
                                        <button type="button" onclick="changeQty('<%=Seq%>',  1)" class="btn-plus">+</button>
									<%ELSE%>
                                        <button type="button" class="btn-minus dis">-</button>
                                        <span class="product-length">1</span>
                                        <button type="button" class="btn-plus dis">+</button>
									<%END IF%>
                                    </div>
                                    <div class="size">사이즈 : <%=SizeCD%></div>
                                    <div class="cost">금액 : <span class="saleprice"><%=FormatNumber(SalePrice,0)%></span>원<%IF SalePriceType = "2" THEN%><span class="employee">임직원가</span><%END IF%></div>

                                    <button type="button" onclick="deleteProduct('<%=Seq%>')" class="btn-hide-selected">삭제</button>
                                </div>
                            </div>
							<%IF EventProdNM <> "" THEN%>
                            <div class="change-cont">
                                <span class="tit">사은품</span><span class="cont"><%=EventProdNM%></span>
                            </div>
							<%END IF%>
                        </li>
<%
'# 1+1 상품
IF OProductCode <> "" AND OSizeCD <> "" THEN
		'# 상품정보
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Product_Select_By_ProductCode"

				.Parameters.Append .CreateParameter("@ProductCode", adInteger, adParamInput, , OProductCode)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				BrandName		= oRs("BrandName")
				ProductName		= oRs("ProductName")
				SalePrice		= oRs("SalePrice")
		ELSE
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||1+1상품은 없는 상품 입니다."
				Response.End
		END IF
		oRs.Close
%>
                        <li class="Product<%=Seq%>">
							<input type="hidden" name="Seq"				value="<%=Seq%>" />
							<input type="hidden" name="ProductCode"		value="<%=OProductCode%>" />
							<input type="hidden" name="SizeCD"			value="<%=OSizeCD%>" />
							<input type="hidden" name="ProductType"		value="O" />
							<input type="hidden" name="SalePriceType"	value="<%=SalePriceType%>" />
							<input type="hidden" name="SalePrice"		value="0" />
                            <div class="selected-cont">
                                <div class="tit"><span class="brand-name"><%=BrandName%></span><span class="item-name"><%=ProductName%></span></div>
                                <div class="cont">
                                    <div class="amount">
                                        <button type="button" class="btn-minus dis">-</button>
                                        <span class="product-length">1</span>
                                        <button type="button" class="btn-plus dis">+</button>
                                    </div>
                                    <div class="size">사이즈 : <%=OSizeCD%></div>
                                    <div class="oneplusone">[1+1상품]</div>
                                </div>
                            </div>
                        </li>
<%
END IF
%>


<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>