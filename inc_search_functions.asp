<%
' Show headlines
Function showNewsHeads(chan,num)
	Dim rsHead, toReturn
	toReturn = ""
	sql = "SELECT TOP "&num&" ID, StoryTitle " _
	& "FROM tblStories " _
	& "WHERE vet = 1 "
	if chan<>0 then
		sql = sql & "AND intchannel="&chan&" "
	end if
	sql = sql & "ORDER BY ID DESC "
	set rsHead = Server.CreateObject("ADODB.Recordset")
	Call rsHead.Open(sql, oDBNews, 3)
	if not rsHead.BOF then
		do while not rsHead.EOF
			toReturn = toReturn & "» <a href="""&siteURL&"/news/news_story.asp?id="&rsHead("ID")&"&channel="&channel
			toReturn = toReturn & "&title="&Server.URLEncode(rsHead("StoryTitle"))
			toReturn = toReturn & """>"&rsHead("StoryTitle")&"</a><br>"&vbCr
			rsHead.MoveNext
		loop
	end if
	showNewsHeads = toReturn
End Function

' Show top 4 stories, then totnum-npics text stories, based on priority, date.
Function showColumnedNews(totnum,sqladd,headcol)
	Dim rsHead, toReturn, tmpCnt, rsTopCat, strTopCat
	toReturn = ""
	' draw container table
	toReturn = toReturn & "<br>"&vbCr
	toReturn = toReturn & "<table border=""0"">" & vbCr
	
	sql = "SELECT DISTINCT TOP "&totnum&" TblStories.ID, StoryTitle, Caption, picture, CAST(StoryIntro AS varchar(300)) AS StoryIntro, priority " _
	& "FROM tblStories " _
	& "JOIN tblStoriesCat ON TblStories.ID = tblStoriesCat.StoryID " _ 
	& "WHERE vet = 1 AND SourceMPID = 0 "
	if sqladd<>"" then
		sql = sql & sqladd
	end if
	'sql = sql & " ORDER BY priority DESC, picture DESC, TblStories.ID DESC "
	sql = sql & " ORDER BY priority DESC, TblStories.ID DESC "
	set rsHead = Server.CreateObject("ADODB.Recordset")
	Call rsHead.Open(sql, oDBNews, 3)
	if not rsHead.BOF then
		tmpCnt = 0
		toReturn = toReturn & "		<tr valign=""top"">" & vbCr
		do while not rsHead.EOF
			sql = "SELECT CatDesc FROM tblCat " & _
			" JOIN tblStoriesCat ON tblStoriesCat.CatID = tblCat.CatID " & _
			" WHERE tblStoriesCat.StoryID = '" & rsHead("id") & "' "
			Set rsTopCat = Server.CreateObject("ADODB.Recordset")
			Call rsTopCat.Open(sql,oDBNews,3)
			if not rsTopCat.bof then
				strTopCat = rsTopCat("CatDesc")
			end if
			
			' draw first column
			toReturn = toReturn & "		<td width=""215"" style=""width:215px"">"
			toReturn = toReturn & "						<span>" & strTopCat & ":</span>"
			toReturn = toReturn & "						<a class=""other-top"" href=""" &siteURL & "/news/news_story.asp?id="&rsHead("id")&"&channel="&channel
			toReturn = toReturn & "&title="&Server.URLEncode(rsHead("StoryTitle"))
			toReturn = toReturn & """>"&rsHead("StoryTitle")&"</a>"
			toReturn = toReturn & "		</td>"
			if tmpCnt = 1 then toReturn = toReturn & "</tr><tr>" ' draw new row
			if tmpCnt = 3 then toReturn = toReturn & "</tr>" ' draw last row
			rsHead.MoveNext
			tmpCnt = tmpCnt + 1
		loop
	end if
	toReturn = toReturn & "</table>" & vbCr
	showColumnedNews = toReturn
End Function

' Show top 4 stories, then totnum-npics text stories, based on priority, date.
Function showColumnedNews2(totnum,sqladd,headcol)
	Dim rsHead, toReturn, tmpCnt, rsTopCat, strTopCat, lastid
	Response.Write  ""
	' draw container table
	Response.Write  "<br>"&vbCr
	Response.Write  "<table border=""0"">" & vbCr
	
	sql = "SELECT DISTINCT TOP "&totnum&" TblStories.ID, StoryTitle, Caption, picture, CAST(StoryIntro AS varchar(300)) AS StoryIntro, priority " _
	& "FROM tblStories " _
	& "JOIN tblStoriesCat ON TblStories.ID = tblStoriesCat.StoryID " _ 
	& "WHERE vet = 1 AND SourceMPID = 0 "
	if sqladd<>"" then
		sql = sql & sqladd
	end if
	'sql = sql & " ORDER BY priority DESC, picture DESC, TblStories.ID DESC "
	sql = sql & " AND embargo <= GETDATE() " ' add embargo (if no embargo, the data will be 01/01/1900 00:00:00)
	sql = sql & " ORDER BY priority DESC, TblStories.ID DESC "
	Response.Write "<!--" & sql & "-->"
	set rsHead = Server.CreateObject("ADODB.Recordset")
	Call rsHead.Open(sql, oDBNews, 3)
	if not rsHead.BOF then
		tmpCnt = 0
		lastid = 0
		Response.Write  "		<tr valign=""top"">" & vbCr
		do while not rsHead.EOF
			'lastid = rsHead("id")
			lastid = lastid & trim(rsHead("id")) & "," ' this adds ID to 'lastid' with comma, turns 'lastid' into a CSV string. 
			' Ideally this should be able to be passed back as a 'WHERE NOT IN  (..) ' clause within the sqladd argument
			sql = "SELECT CatDesc FROM tblCat " & _
			" JOIN tblStoriesCat ON tblStoriesCat.CatID = tblCat.CatID " & _
			" WHERE tblStoriesCat.StoryID = '" & rsHead("id") & "' "
			Set rsTopCat = Server.CreateObject("ADODB.Recordset")
			Call rsTopCat.Open(sql,oDBNews,3)
			if not rsTopCat.bof then
				strTopCat = rsTopCat("CatDesc")
			end if
			
			' draw first column
			Response.Write  "		<td width=""215"" style=""width:215px"">"
			Response.Write  "						<span>" & strTopCat & ":</span>"
			Response.Write  "						<a class=""other-top"" href=""" &siteURL & "/news/news_story.asp?id="&rsHead("id")&"&channel="&channel
			Response.Write  "&title="&Server.URLEncode(rsHead("StoryTitle"))
			Response.Write  """>"&rsHead("StoryTitle")&"</a>"
			Response.Write  "		</td>"
			if tmpCnt = 1 then Response.Write  "</tr><tr>" ' draw new row
			if tmpCnt = 3 then Response.Write  "</tr>" ' draw last row
			rsHead.MoveNext
			tmpCnt = tmpCnt + 1
		loop
	end if
	' lastid might end up a bit messy. Lets clean it up before returning it.
	If Right(lastid,1) = "," then lastid = Left(lastid,Len(lastid)-1)
	Response.Write  "</table>" & vbCr
	toReturn = lastid
	showColumnedNews2 = toReturn
End Function

Function showTopStory(chan,catid)
	' accepts cat based for cat homepages, if 0 then don't filter etc
	' write out top story for home page, and return id pulled out
	dim toReturn, lastid, rsHead, txt, storyTitle, storyIntro, storyID
	oDBNews.Execute("SET DATEFORMAT dmy")
	sql = "SELECT DISTINCT TOP 1 TblStories.ID, " _ 
	& "StoryTitle, Caption, picture, CAST(StoryIntro AS varchar(300)) AS Intro1, CAST(StoryBody AS varchar(120)) AS Intro2, priority " _
	& "FROM tblStories " _
	' cat based ?
	if catid>0 then 
		sql = sql & "JOIN tblStoriesCat ON TblStories.ID = tblStoriesCat.StoryID " 
		sql = sql & "WHERE tblStoriesCat.CatID = '" & catid & "' "
		sql = sql & "AND vet = 1 AND SourceMPID = 0 "
	else	
		sql = sql & "WHERE vet = 1 AND SourceMPID = 0 "
	end if
	if chan<>"" then sql = sql & " AND intChannel = '" & chan & "' "
	sql = sql & " AND embargo <= GETDATE() " ' add embargo (if no embargo, the data will be 01/01/1900 00:00:00)
	'sql = sql & " ORDER BY priority DESC, picture DESC, TblStories.ID DESC "
	sql = sql & " ORDER BY priority DESC, TblStories.ID DESC "
	'Response.Write sql	
	set rsHead = Server.CreateObject("ADODB.Recordset")
	Call rsHead.Open(sql, oDBNews, 3)
	if not rsHead.BOF then
		storyTitle = rsHead("StoryTitle")
		storyID = rsHead("id")
		if rsHead("Intro1")<>"" then
			storyIntro = rsHead("Intro1") ' our intro
		else
			storyIntro = rsHead("Intro2") ' UW guys didn't put in an intro, so we would have to go for the first 300 chars of the story body instead!
		end if
		lastid = rsHead("id") 
		txt = txt & "			<h2 class=""leadstory""><a title=""View the story '" & storyTitle & "'"" href=""" & siteURL & "/news/news_story.asp?id=" & storyID & "&title=" &Server.URLEncode(storyTitle) & """>" &  storyTitle & "</a></h2>" & vbCr
		txt = txt & "			<img src=""" & siteURL & "/news/images/" & storyID & ".jpg"" class=""leadstory-img"" width=""150"">" & vbCr
		txt = txt & "			<div class=""leadstory-intro"">" & storyIntro
		txt = txt & "				<a class=""read-more"" title=""View the story '" & storyTitle & "'"" href=""" & siteURL & "/news/news_story.asp?id=" & storyID & "&title=" &Server.URLEncode(storyTitle) & """>...more&nbsp;»</a>" & vbCr	
		txt = txt & "			</div>" & vbCr
	end if
	Response.Write txt
	showTopStory = lastid 'return the ID so we know to not show it in later 'showPriNews' calls
End Function

Function showSponsoredArticle(chan,articleid)
	' same as above, just does it a bit more snazzy way...
	' accepts cat based for cat homepages, if 0 then don't filter etc
	' write out top story for home page, and return id pulled out
	dim toReturn, lastid, rsHead, txt, storyTitle, storyIntro, storyID
	oDBNews.Execute("SET DATEFORMAT dmy")
	sql = "SELECT DISTINCT TOP 1 TblStories.ID, " _ 
	& "StoryTitle, Caption, picture, CAST(StoryIntro AS varchar(300)) AS Intro1, CAST(StoryBody AS varchar(120)) AS Intro2, priority " _
	& "FROM tblStories " _
	' cat based ?
	if articleid>0 then 
		sql = sql & "WHERE TblStories.ID = '" & articleid & "' "
		sql = sql & "AND vet = 1 AND SourceMPID = 0 "
	else	
		sql = sql & "WHERE vet = 1 AND SourceMPID = 0 "
	end if
	if chan<>"" then sql = sql & " AND intChannel = '" & chan & "' "
	sql = sql & " AND embargo <= GETDATE() " ' add embargo (if no embargo, the data will be 01/01/1900 00:00:00)
	'sql = sql & " ORDER BY priority DESC, picture DESC, TblStories.ID DESC "
	sql = sql & " ORDER BY priority DESC, TblStories.ID DESC "
	'Response.Write sql
	set rsHead = Server.CreateObject("ADODB.Recordset")
	Call rsHead.Open(sql, oDBNews, 3)
	if not rsHead.BOF then
		storyTitle = rsHead("StoryTitle")
		storyID = rsHead("id")
		if rsHead("Intro1")<>"" then
			storyIntro = rsHead("Intro1") ' our intro
		else
			storyIntro = rsHead("Intro2") ' UW guys didn't put in an intro, so we would have to go for the first 300 chars of the story body instead!
		end if
		lastid = rsHead("id") 
		txt = txt & "<h2 class=""leadstory"">"
		txt = txt & "	<a title=""View the story '" & storyTitle & "'"" href=""" & siteURL & "/news/news_story.asp?id=" & storyID & "&title=" &Server.URLEncode(storyTitle) & """>" &  storyTitle & "</a>"
		txt = txt & "</h2>" & vbCr
		txt = txt & "<a href=""" & siteURL & "/news/news_story.asp?id=" & storyID & "&title=" &Server.URLEncode(storyTitle) & """><img src=""" & siteURL & "/news/images/" & storyID & ".jpg"" alt=""" & storyTitle & """ class=""leadstory-img"" width=""150"">" & vbCr
		txt = txt & "<div class=""leadstory-intro"">" & storyIntro
		txt = txt & "	<a class=""read-more"" title=""View the story '" & storyTitle & "'"" href=""" & siteURL & "/news/news_story.asp?id=" & storyID & "&title=" &Server.URLEncode(storyTitle) & """>...more&nbsp;»</a>" & vbCr	
		txt = txt & "</div>" & vbCr
	end if
	Response.Write txt
	showSponsoredArticle = lastid 'return the ID so we know to not show it in later 'showPriNews' calls
End Function

' Show top npics picture stories, then totnum-npics text stories, based on priority, date.
Function showPriNews(style,npics,totnum,sqladd,headcol)
	Dim rsHead, toReturn, tmpCnt, storyID, storyTitle, storyIntro, rsTopCat, strTopCat, storyAuthor
	toReturn = ""
	if style > 0 then toReturn = toReturn & "<ol class=""post-list"">"
	sql = "SELECT DISTINCT TOP "&totnum&" TblStories.ID, " _ 
	& "StoryTitle, Caption, picture, SourceName, CAST(StoryIntro AS varchar(300)) AS Intro1, CAST(StoryBody AS varchar(120)) AS Intro2, priority " _
	& "FROM tblStories " _
	& "JOIN tblStoriesCat ON TblStories.ID = tblStoriesCat.StoryID " _ 
	& "WHERE vet = 1 AND SourceMPID = 0 "
	if sqladd<>"" then
		sql = sql & sqladd
	end if
	'sql = sql & " ORDER BY priority DESC, picture DESC, TblStories.ID DESC "
	sql = sql & " AND embargo <= GETDATE() " ' add embargo (if no embargo, the data will be 01/01/1900 00:00:00)
	sql = sql & " ORDER BY priority DESC, TblStories.ID DESC "
	Response.Write sql & "<br>"
	set rsHead = Server.CreateObject("ADODB.Recordset")
	Call rsHead.Open(sql, oDBNews, 3)
	if not rsHead.BOF then
		tmpCnt = 0
		do while not rsHead.EOF
			'Response.Write "got here"
			storyTitle = rsHead("StoryTitle")
			storyID = rsHead("id")
			if rsHead("Intro1")<>"" then
				storyIntro = rsHead("Intro1") ' our intro
			else
				storyIntro = rsHead("Intro2") ' UW guys didn't put in an intro, so we would have to go for the first 300 chars of the story body instead!
			end if
			storyAuthor = get_author(rsHead("SourceName"))
			sql = "SELECT CatDesc FROM tblCat " & _
			" JOIN tblStoriesCat ON tblStoriesCat.CatID = tblCat.CatID " & _
			" WHERE tblStoriesCat.StoryID = '" & storyID & "' "
			'Response.Write sql
			Set rsTopCat = Server.CreateObject("ADODB.Recordset")
			Call rsTopCat.Open(sql,oDBNews,3)
			if not rsTopCat.bof then
				strTopCat = rsTopCat("CatDesc")
			end if
			if style=1 then ' showing on an index ie. home page / news ?
				toReturn = toReturn & "<li class=""clearfix"">" & vbCr
				toReturn = toReturn & "	<h3><a href=""" & siteURL & "/news/news_story.asp?id=" & storyID & "&title=" & Server.URLEncode(storyTitle) & """>" 
				toReturn = toReturn & storyTitle & "</a></h3>" & vbCr
				toReturn = toReturn & "		<span class=""post-meta"">" & strTopCat & " | " & storyAuthor & "</span><br>" & vbCr
				if rsHead("picture")="picture!" then
					toReturn = toReturn & " 		<div class=""thumbimg"">" & vbCr
					toReturn = toReturn & " 			<img src=""" &siteURL & "/news/images/"& storyID &".jpg"">" & vbCr
					toReturn = toReturn & "			</div>" & vbCr
				end if
				toReturn = toReturn & "			<p class=""post-excerpt"">" & storyIntro & "..</p>"
				toReturn = toReturn & "				<a class=""read-more"" style=""display:inline;"" "
				toReturn = toReturn & 					"	title=""View the entry '" & storyTitle & "'"" "
				toReturn = toReturn & 					"	href=""" & siteURL & "/news/news_story.asp?id=" & storyID & "&title=" & Server.URLEncode(storyTitle) & """>"
				toReturn = toReturn & 					"Read more&nbsp;»</a>"
				'toReturn = toReturn & "			</p>"
				toReturn = toReturn & "</li>"
			end if
			if style=2 then
				toReturn = toReturn & "<li class=""clearfix"">" & vbCr
				toReturn = toReturn & "	<h3><a href=""" & siteURL & "/news/news_story.asp?id=" & storyID & "&title=" & Server.URLEncode(storyTitle) & """>" 
				toReturn = toReturn & storyTitle & "</a></h3>" & vbCr
				toReturn = toReturn & "		<span class=""post-meta"">" & strTopCat & " | " & storyAuthor & "</span><br>" & vbCr
				'toReturn = toReturn & " 		<div class=""thumbimg"">" & vbCr
				'toReturn = toReturn & " 			<img src=""" &siteURL & "/news/images/"& storyID &".jpg"">" & vbCr
				'toReturn = toReturn & "			</div>" & vbCr
				toReturn = toReturn & "			<p class=""post-excerpt"">" & Left(storyIntro,100) & "..</p>"
				toReturn = toReturn & "				<a class=""read-more"" style=""display:inline;"" "
				toReturn = toReturn & 					"	title=""View the entry '" & storyTitle & "'"" "
				toReturn = toReturn & 					"	href=""" & siteURL & "/news/news_story.asp?id=" & storyID & "&title=" & Server.URLEncode(storyTitle) & """>"
				toReturn = toReturn & 					"Read more&nbsp;»</a>"
				'toReturn = toReturn & "			</p>"
				toReturn = toReturn & "</li>"
				'if tmpCnt<npics then
				'	toReturn = toReturn & "<div class=""" & headcol & """><h2>"
				'else 
				'	toReturn = toReturn & "» "
				'end if
				'toReturn = toReturn & "<a href=""" &siteURL & "/news/news_story.asp?id="&rsHead("id")&"&channel="&channel
				'toReturn = toReturn & "&title="&Server.URLEncode(rsHead("StoryTitle"))
				'toReturn = toReturn & """>"&rsHead("StoryTitle")&"</a>"
				'if tmpCnt<npics then
				'	toReturn = toReturn & "</h2></div>"&vbCr
				'end if
				'if rsHead("picture") = "picture!" and tmpCnt<npics then
				'	toReturn = toReturn & "<img src=""" &siteURL & "/news/images/"&rsHead("id")&".jpg"" alt="""&rsHead("Caption")&""" width=""140"" height=""100"" class=""news_img_right"">"
				'end if

				'	toReturn = toReturn & rsHead("StoryIntro") _ 
				'	& "<br><div class=""" & headcol & """>» <a href=""" &siteURL & "/news/news_story.asp?id="&rsHead("id")&"&channel="&channel _
				'	& "&title="&Server.URLEncode(rsHead("StoryTitle")) _ 
				'	& """>Read more</a></div>" _
				'	& "<div class=""clr""></div>" _
				'	& "<div class=""br2""></div>"&vbCr
				'else
				'	toReturn = toReturn & "<br>"&vbCr
				'end if
			end if
			rsHead.MoveNext
			tmpCnt = tmpCnt + 1
		loop
	end if
	if style > 0 then toReturn = toReturn & "</ol>"
	showPriNews = toReturn
End Function

Function showEvents(num)
	Dim rsHead, toReturn, tmpCount
	toReturn = ""
	sql = "SELECT TOP "&num&" ID, Title, dt, location, url " _
	& "FROM Events " _
	& "WHERE vet = 1 "
	'if num<=10 then ' AG commented out, no point in not showing upcoming events if you're drawing out more than 10 eg. from see all events...
		sql = sql & "AND dt > DATEADD(dd,-1,GETDATE())  "
	'end if
	sql = sql & "ORDER BY dt ASC " ' closest at top
	set rsHead = Server.CreateObject("ADODB.Recordset")
	Call rsHead.Open(sql, oDBNews, 3)
	if not rsHead.BOF then
		toReturn = toReturn & "<ol class=""post-list"">"
		tmpCount = 0
		do while not rsHead.EOF
			'toReturn = toReturn & "» <a href="""&siteURL&"/events/event.asp?id="&rsHead("ID")&"&channel="&channel
			'toReturn = toReturn & "&title="&Server.URLEncode(rsHead("Title"))
			'toReturn = toReturn & """>"&rsHead("Title")&"</a><br>"&vbCr
			if tmpCount = 1 then ' have we reached last row ?
				toReturn = toReturn & "	<li class=""clearfix last"">" & vbCr
			else
				toReturn = toReturn & "	<li class=""clearfix"">" & vbCr
			end if
			toReturn = toReturn & "		<h3 class=""event-title""><a href=""" & siteURL & "/events/view_event.asp?id=" & rsHead("ID") & """>" & rsHead("Title") & "</a></h3>"
			toReturn = toReturn & "<span class=""post-meta"">" & rsHead("location") & ", " & rsHead("dt") & "</span>"
			toReturn = toReturn &	"	</li>"
			tmpCount = tmpCount + 1
			rsHead.MoveNext
		loop
		toReturn = toReturn &	"</ol>" & vbCr
	end if
	
	showEvents = toReturn
End Function

function get_author(num)
	dim toReturn, rsTmp
	if num <> "Utility Week" then
		sql = "SELECT name, email FROM authors WHERE authorid='" & num & "'"
		Set rsTmp = Server.CreateObject("ADODB.Recordset")
		Call rsTmp.Open(sql,oDBNews,3)
		if not rsTmp.bof then
			toReturn = "<a href=""mailto:" & rsTmp("email") & """>" & rsTmp("name") & "</a>"
		end if
		rsTmp.Close
		Set rsTmp = Nothing
	else ' sometimes it's just 'Utility Week' !
		toReturn = num
	end if
	get_author = toReturn
end function

function get_blogs(typ,num)
	' typ = 1: Connected, 0: Disconnector
	dim toReturn, rsTmp, blogid, title, tmpCount
	toReturn = ""
	sql = "SELECT TOP " & num & " blogid, title FROM blog WHERE channel = '" & typ & "' ORDER BY blogid DESC "
	Set rsTmp = Server.CreateObject("ADODB.Recordset")
	Call rsTmp.Open(sql,oDBNews,3)
	if not rsTmp.bof then
		tmpCount = 0
		toReturn = toReturn & "<ol class=""post-list"">" & vbCr
		do while not rsTmp.eof
			blogid = rsTmp("blogid")
			title = rsTmp("title")
			if tmpCount = num - 1 then
				toReturn = toReturn & "<li class=""clearfix last"">" & vbCr
			else
				toReturn = toReturn & "<li class=""clearfix"">" & vbCr
			end if
			toReturn = toReturn & "			<p><a title=""" & vbCr
			toReturn = toReturn & title & """ href=""" & siteURL 
			toReturn = toReturn & "/blog/view_entry.asp?id=" & blogid &"&channel=" & channel & """>" & title & "</a>" & vbCr
			toReturn = toReturn & "			</p>" & vbCr
			toReturn = toReturn & "</li>" & vbCr
			tmpCount = tmpCount + 1
			rsTmp.MoveNext()
		Loop
		toReturn = toReturn & "</ol>" & vbCr
	end if
	rsTmp.Close
	Set rsTmp = Nothing
	get_blogs = toReturn
end function

Function getArtSubStatus()
	'get article subscription status from currently logged in user
	dim toReturn, rsTmp
	if userID>0 then 
		sql = "SELECT dtSubscriptionExpiry FROM tblUsers WHERE userid = '" & inc_sql_escape(userid) & "' AND dtSubscriptionExpiry>=GETDATE() "
		'sql = "SELECT articleSubStatus FROM tblUsers WHERE UserID = '" & inc_sql_escape(Session("userID")) & "'"
		Set rsTmp = Server.CreateObject("ADODB.Recordset")
		Call rsTmp.Open(sql,oDBUsers,3)
		if not rsTmp.bof then
			toReturn = True
		end if
		rsTmp.Close
		Set rsTmp = Nothing
	else
		toReturn = false
	end if
	getArtSubStatus = toReturn
end Function
%>