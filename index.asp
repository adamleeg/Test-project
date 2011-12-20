<% Option Explicit %>
<!-- #include virtual="/../includes/globals.asp" -->
<!-- #include virtual="/../includes/inc_db_connections.asp" -->
<!-- #include virtual="/../includes/inc_general_functions.asp" -->
<!-- #include virtual="/../includes/inc_search_functions.asp" -->
<!-- #include virtual="/../includes/jobs/functions.asp" -->
<!-- #include virtual="/../includes/jobs/include_login.asp" -->
<%
	db_connect_all()	' Connect to all dbs
	persist_load_userID()	' Important for job centre
	get_user_details()	
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	section = "home"
	dim lastid
	showKP = TRUE ' show Knowledge Partners
%>
<!-- #include virtual="/../includes/nav_header.asp" -->

<!-- Page Wrapper -->
<div id="wrapper">
<!-- Main body -->
<div id="main"><div id="content">

<!-- Nav Top Page -->
<!-- #include virtual="/../includes/nav_top_default.asp" -->

<div class="br2"></div>
<!-- content container -->
<table border="0" cellspacing="0" cellpadding="0" width="100%">
	<tr valign="top">
		<td>
			<!-- top box - latest news, subscribe box, jobs -->
			<table border="0" width="650" style="width:650px;">
				<tr valign="top">
					<td>
						<!-- headline box -->
						<div class="home_headline_box">
					
							<!-- lead story -->
							<%
							'<h2 class="leadstory">XXLEAD STORY TITLEXX</h2>
							'<img src="http://www.utilityweek.co.uk/news/05%20Bills%20man%20worried%20%28IS%29THMB.jpg" class="leadstory-img">
							'<div class="leadstory-intro">xxx
							'<a class="read-more" title="xxx" href="">...more&nbsp;»</a>	
							'</div>
							%>
							<% lastid = showTopStory(0,0) %>
							<!-- lead story end -->
					
							<div class="br2"></div>
							<!-- other 4 top stories -->
							<h2 class="top-stories">Other top stories</h2>
							<% 
							lastid = showColumnedNews2(4,"AND intChannel=0 and TblStories.ID<>'" & lastid & "' ","red_head2") 
							%>
							<!-- other 4 top stories end -->
						
						<!-- headline box end -->
						</div>
					</td>
					<td>
						<!-- second col within top box -->
						<div style="width:200px;margin-left:5px;margin-right:5px;">
							<!-- magazine subscriptions ad box -->
							<div class="mag_sub_box">
					
								<a href="http://www.fhgmedia.com/shop/shopdisplayproducts.asp?id=25&cat=Utility+Week">
									<img class="mag_sub" alt="Subscribe to Utility Week magazine" src="http://www.utilityweek.co.uk/images/cover.jpg" width="80" height="100">
								</a>
								<h2 class="title">FOR THE <br>FULL<br> STORY...</h2> <BR>
								<a href="http://www.fhgmedia.com/shop/shopdisplayproducts.asp?id=25&cat=Utility+Week">subscribe <br>now</a>»
								<div class="br2"></div>
							</div>
							<!-- magazine subscriptions ad box END -->
							
							<!-- jobs content -->
							<div id="jobs">
								<h2><a href="/jobs/">TOP JOBS</a></h2>
								<%=get_job_headlines(99,0,5)%>
							</div>
							<!-- jobs content END -->
							
							<!-- spacer, to match start of news boxes -->
							<div style="line-height:31px;">&nbsp;</div>
							<!-- spacer, to match start of news boxes -->
			
						</div>
						<!-- second col within top box END -->
					</td>
				</tr>
			</table>
			<!-- top box - latest news, subscribe box, jobs END -->
			<!-- second row latest news / current features-->
			<table border="0" width="650" style="width:650px;">
				<!--<div class="br2"></div>-->
				<!-- container for latest news / current features boxes -->
				<tr valign="top">
					<td width="42%">
							<h2>Latest News</h2>
							<br>
							<%
							dim sqlStr : sqlStr = "" ' init var and blank it
							if InStr(lastid,",")>0 then ' test for a comma existing in this string. If so, that means it's a CSV.
								' Make an SQL statement with a NOT IN (...) clause
								sqlStr = "AND intChannel='0' AND tblStories.ID NOT IN (" & lastid & ") "
							else ' Not a CSV ? Change the SQL statement, then
								sqlStr = "AND intChannel='0' tblStories.ID < '" & lastid & "' "
							end if	
							Response.Write showPriNews(1,2,5,sqlStr,"red_head2") %>
							<div class="see-all"><a href="<%=siteURL%>/news">See all news&nbsp;»</a></div>
							<!-- latest news END -->
					</td>
					<td width="2%">&nbsp;</td>
					<td id="dot-divider" width="1%">&nbsp;</td>
					<td width="2%">&nbsp;</td>
					<td width="42%">
							<!-- current features  -->
							<h2>Current Features</h2>
							<br>
							<% Response.Write showPriNews(1,3,3,"AND intChannel=1","red_head2") %>
							<div class="see-all"><a href="<%=siteURL%>/features">See all features&nbsp;»</a></div>
							<br>
							<% ' Homepage MPU Zone 418
							adRegion = 418
							adHeight = 250
							adWidth = 300
							adSlot  = 1
							adText = "Visit our Sponsors"
							%>
							<!-- #include virtual="/../includes/include_fhg_advert.asp" -->	
					</td>
					<!-- current features END -->
			</tr>
		</table>
			<!-- second row latest news / current features end -->
		<!-- container for  boxes END -->
	</td>
	<td>
		<!-- column 2 - blog, MPU, events, talk boards-->
		<!-- #include virtual="/../includes/nav_right_default.asp" -->
	</td>
</tr>
</table>
<!-- content container END -->
<div class="br2"></div>
<!-- #include virtual="/../includes/nav_footer.asp" -->

<!-- Nav Top Page END -->
<% 
	db_disconnect_all()	' Disconnect from all dbs
%>

</div></div>
<!-- Main body End-->
</div>
<!-- Page Wrapper end -->

</body>
</html>
