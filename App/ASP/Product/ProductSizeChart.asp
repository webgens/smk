<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'ProductSizeChart.asp - 상품 사이즈 정보
'Date		: 2019.01.19
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
PageCode1 = "00"
PageCode2 = "00"
PageCode3 = "00"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

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
%>


<!-- #include virtual="/INC/Header.asp" -->
	<style type="text/css">
		.ofh { overflow: hidden !important; }
		.detail-explanation .img-all img { width: 100%; }
		.selected-cont .cont .cost { display: inline-block; margin-left: 11px; font-size: 11px; color: #b4b4b4; }
		.selected-cont .cont .cost .employee { color: #e62019; }
		.selected-cont .cont .oneplusone { display: inline-block; margin-left: 11px; font-size: 11px; color: #e62019; }
		.onePlus-item-list { padding-top: 0; border-top: none; }
		.onePlus-item-list .inform .cont .price { display: block; font-size: 11px; }
		.onePlus-item-list .inform .cont .price>em { font-size: 14px; font-weight: 800; }
	</style>

<!-- #include virtual="/INC/PopupTop.asp" -->


    <section class="wrap-pop" id="pop-SizeChart" style="display:block">
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="full">
                <div class="tit-pop">
                    <div class="tit">사이즈 차트</div>
                    <button class="btn-hide-pop" onclick="APP_PopupHistoryBack();">닫기</button>
                </div>

                <div class="container-pop">
                    <div class="contents">
                        <div class="wrap-sizeChart">
                            <div id="tabs1" class="tab" data-use="">
                                <ul class="tab-selector">
                                    <li class="active part-2"><a href="javascript:;" data-target="tabs-col11">ADULT</a></li>
                                    <li class="part-2"><a href="javascript:;" data-target="tabs-col21">KIDS</a></li>
                                </ul>
                                <div id="tabs-col11" class="tab-panel">
                                    <div id="sizeChart_men">
                                        <div class="pop-accordion-selector">
                                            <button type="button" class="tit" data-target="sizeChart_men">MEN</button>
                                        </div>
                                        <div class="pop-accordion-panel">
                                            <div class="cont">
                                                <table class="table-hoz">
                                                    <caption class="hidden">한국, 미국, 유럽의 성인 남자 신발 사이즈 비교표</caption>
                                                    <colgroup>
                                                        <col>
                                                        <col style="width: 33.333333%">
                                                        <col style="width: 33.333333%">
                                                    </colgroup>

                                                    <thead>
                                                        <tr>
                                                            <th scope="col">한국</th>
                                                            <th scope="col">미국</th>
                                                            <th scope="col">유럽</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        <tr>
                                                            <td>245</td>
                                                            <td>6.5</td>
                                                            <td>40</td>
                                                        </tr>
                                                        <tr>
                                                            <td>250</td>
                                                            <td>7</td>
                                                            <td>40.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>255</td>
                                                            <td>7.5</td>
                                                            <td>41</td>
                                                        </tr>
                                                        <tr>
                                                            <td>260</td>
                                                            <td>8</td>
                                                            <td>41.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>265</td>
                                                            <td>8.5</td>
                                                            <td>42</td>
                                                        </tr>
                                                        <tr>
                                                            <td>270</td>
                                                            <td>9.5</td>
                                                            <td>42.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>275</td>
                                                            <td>9.5</td>
                                                            <td>43</td>
                                                        </tr>
                                                        <tr>
                                                            <td>280</td>
                                                            <td>10</td>
                                                            <td>43.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>285</td>
                                                            <td>10.5</td>
                                                            <td>44</td>
                                                        </tr>
                                                        <tr>
                                                            <td>290</td>
                                                            <td>11</td>
                                                            <td>44.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>295</td>
                                                            <td>11.5</td>
                                                            <td>45</td>
                                                        </tr>
                                                        <tr>
                                                            <td>300</td>
                                                            <td>12</td>
                                                            <td>45.5</td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </div>
                                        </div>
                                    </div>

                                    <div id="sizeChart_women">
                                        <div class="pop-accordion-selector">
                                            <button type="button" class="tit" data-target="sizeChart_women">WOMEN</button>
                                        </div>
                                        <div class="pop-accordion-panel">
                                            <div class="cont">
                                                <table class="table-hoz">
                                                    <caption class="hidden">한국, 미국, 유럽의 성인 여자 신발 사이즈 비교표</caption>
                                                    <colgroup>
                                                        <col>
                                                        <col style="width: 33.333333%">
                                                        <col style="width: 33.333333%">
                                                    </colgroup>

                                                    <thead>
                                                        <tr>
                                                            <th scope="col">한국</th>
                                                            <th scope="col">미국</th>
                                                            <th scope="col">유럽</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        <tr>
                                                            <td>220</td>
                                                            <td>5</td>
                                                            <td>35</td>
                                                        </tr>
                                                        <tr>
                                                            <td>225</td>
                                                            <td>5.5</td>
                                                            <td>35.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>230</td>
                                                            <td>6</td>
                                                            <td>36</td>
                                                        </tr>
                                                        <tr>
                                                            <td>235</td>
                                                            <td>6.5</td>
                                                            <td>36.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>240</td>
                                                            <td>7</td>
                                                            <td>37</td>
                                                        </tr>
                                                        <tr>
                                                            <td>245</td>
                                                            <td>7.5</td>
                                                            <td>37.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>250</td>
                                                            <td>8</td>
                                                            <td>38</td>
                                                        </tr>
                                                        <tr>
                                                            <td>255</td>
                                                            <td>8.5</td>
                                                            <td>38.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>260</td>
                                                            <td>9</td>
                                                            <td>39</td>
                                                        </tr>
                                                        <tr>
                                                            <td>265</td>
                                                            <td>9.5</td>
                                                            <td>39.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>270</td>
                                                            <td>10</td>
                                                            <td>40</td>
                                                        </tr>
                                                        <tr>
                                                            <td>275</td>
                                                            <td>10.5</td>
                                                            <td>40.5</td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div id="tabs-col21" class="tab-panel">
                                    <div id="sizeChart_kid1">
                                        <div class="pop-accordion-selector">
                                            <button type="button" class="tit" data-target="sizeChart_kid1">3~36개월</button>
                                        </div>
                                        <div class="pop-accordion-panel">
                                            <div class="cont">
                                                <table class="table-hoz">
                                                    <caption class="hidden">한국, 미국, 유럽의 3~36개월 아이 신발 사이즈 비교표</caption>
                                                    <colgroup>
                                                        <col>
                                                        <col style="width: 33.333333%">
                                                        <col style="width: 33.333333%">
                                                    </colgroup>

                                                    <thead>
                                                        <tr>
                                                            <th scope="col">한국</th>
                                                            <th scope="col">미국</th>
                                                            <th scope="col">유럽</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        <tr>
                                                            <td>80</td>
                                                            <td>2</td>
                                                            <td>17</td>
                                                        </tr>
                                                        <tr>
                                                            <td>85</td>
                                                            <td>2.5</td>
                                                            <td>18</td>
                                                        </tr>
                                                        <tr>
                                                            <td>90</td>
                                                            <td>3</td>
                                                            <td>18.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>95</td>
                                                            <td>3.5</td>
                                                            <td>19</td>
                                                        </tr>
                                                        <tr>
                                                            <td>100</td>
                                                            <td>4</td>
                                                            <td>19.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>105</td>
                                                            <td>4.5</td>
                                                            <td>20</td>
                                                        </tr>
                                                        <tr>
                                                            <td>110</td>
                                                            <td>5</td>
                                                            <td>21</td>
                                                        </tr>
                                                        <tr>
                                                            <td>115</td>
                                                            <td>5.5</td>
                                                            <td>21.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>120</td>
                                                            <td>6</td>
                                                            <td>22</td>
                                                        </tr>
                                                        <tr>
                                                            <td>125</td>
                                                            <td>6.5</td>
                                                            <td>22.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>130</td>
                                                            <td>7</td>
                                                            <td>23.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>135</td>
                                                            <td>7.5</td>
                                                            <td>24</td>
                                                        </tr>
                                                        <tr>
                                                            <td>140</td>
                                                            <td>8</td>
                                                            <td>25</td>
                                                        </tr>
                                                        <tr>
                                                            <td>145</td>
                                                            <td>8.5</td>
                                                            <td>25.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>150</td>
                                                            <td>9</td>
                                                            <td>26</td>
                                                        </tr>
                                                        <tr>
                                                            <td>155</td>
                                                            <td>9.5</td>
                                                            <td>26.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>160</td>
                                                            <td>10</td>
                                                            <td>27</td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </div>
                                        </div>
                                    </div>
                                    <div id="sizeChart_kid2">
                                        <div class="pop-accordion-selector">
                                            <button type="button" class="tit" data-target="sizeChart_kid2">4~7세</button>
                                        </div>
                                        <div class="pop-accordion-panel">
                                            <div class="cont">
                                                <table class="table-hoz">
                                                    <caption class="hidden">한국, 미국, 유럽의 4~7세 아이 신발 사이즈 비교표</caption>
                                                    <colgroup>
                                                        <col>
                                                        <col style="width: 33.333333%">
                                                        <col style="width: 33.333333%">
                                                    </colgroup>

                                                    <thead>
                                                        <tr>
                                                            <th scope="col">한국</th>
                                                            <th scope="col">미국</th>
                                                            <th scope="col">유럽</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        <tr>
                                                            <td>165</td>
                                                            <td>10.5C</td>
                                                            <td>27.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>170</td>
                                                            <td>11</td>
                                                            <td>28</td>
                                                        </tr>
                                                        <tr>
                                                            <td>175</td>
                                                            <td>11.5C</td>
                                                            <td>28.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>180</td>
                                                            <td>12C</td>
                                                            <td>29.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>185</td>
                                                            <td>12.5C</td>
                                                            <td>30</td>
                                                        </tr>
                                                        <tr>
                                                            <td>190</td>
                                                            <td>13C</td>
                                                            <td>31</td>
                                                        </tr>
                                                        <tr>
                                                            <td>195</td>
                                                            <td>13.5C</td>
                                                            <td>31.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>200</td>
                                                            <td>1Y</td>
                                                            <td>32</td>
                                                        </tr>
                                                        <tr>
                                                            <td>205</td>
                                                            <td>1.5Y</td>
                                                            <td>33</td>
                                                        </tr>
                                                        <tr>
                                                            <td>210</td>
                                                            <td>2Y</td>
                                                            <td>33.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>215</td>
                                                            <td>2.5Y</td>
                                                            <td>34</td>
                                                        </tr>
                                                        <tr>
                                                            <td>220</td>
                                                            <td>3Y</td>
                                                            <td>35</td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </div>
                                        </div>
                                    </div>
                                    <div id="sizeChart_kid3">
                                        <div class="pop-accordion-selector">
                                            <button type="button" class="tit" data-target="sizeChart_kid3">8~13세</button>
                                        </div>
                                        <div class="pop-accordion-panel">
                                            <div class="cont">
                                                <table class="table-hoz">
                                                    <caption class="hidden">한국, 미국, 유럽의 8~13세 아이 신발 사이즈 비교표</caption>
                                                    <colgroup>
                                                        <col>
                                                        <col style="width: 33.333333%">
                                                        <col style="width: 33.333333%">
                                                    </colgroup>

                                                    <thead>
                                                        <tr>
                                                            <th scope="col">한국</th>
                                                            <th scope="col">미국</th>
                                                            <th scope="col">유럽</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        <tr>
                                                            <td>225</td>
                                                            <td>3.5Y</td>
                                                            <td>35.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>230</td>
                                                            <td>4Y</td>
                                                            <td>36</td>
                                                        </tr>
                                                        <tr>
                                                            <td>235</td>
                                                            <td>4.5Y</td>
                                                            <td>26.5</td>
                                                        </tr>
                                                        <tr>
                                                            <td>240</td>
                                                            <td>5Y</td>
                                                            <td>37</td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </section>


<!-- #include virtual="/INC/FooterNone.asp" -->
<!-- #include virtual="/INC/PopupBottom.asp" -->