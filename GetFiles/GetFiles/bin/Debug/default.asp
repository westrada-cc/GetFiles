<!-- #include virtual="initial_default.asp" -->
<!--#include virtual="ASPModules/AdvancedFeaturedProperty.asp"-->
<%
TemplatePath = "/GlobalTemplates/AgentTemplates/RM1018-RRE-EXECUTIVES/"
session("TemplatePath") = TemplatePath
%>
<%

SUB DisplayRSSFeed2(URLToRSS, MaxNumberOfItems)
	
	' =========== configuration =====================
	' Item template.
	' {LINK} will be replaced with item link
	' {TITLE} will be replaced with item title
	' {DESCRIPTION} will be replaced with item description
	' {CREATOR} will be replaced with item Creator
	'ItemTemplate = "<tr><td><a target='_blank' href=" & """{LINK}""" & ">{TITLE}</a><BR>{DESCRIPTION}</td></tr>"
	
	ItemTemplate = ""
	ItemTemplate = ItemTemplate & "<div class=""rre_blog_entry"">"
	ItemTemplate = ItemTemplate & "<div class=""rre_blog_date"">{DATE}</div>"
	ItemTemplate = ItemTemplate & "<div class=""rre_blog_creator"">Created by: {CREATOR}</div>"
	ItemTemplate = ItemTemplate & "<div class=""rre_blog_title""><a href=""{LINK}"" target=""_blank"">{TITLE}</a></div>"
	ItemTemplate = ItemTemplate & "	<div class=""rre_blog_text"">"
	ItemTemplate = ItemTemplate & "	{DESCRIPTION}"
	ItemTemplate = ItemTemplate & "	</div>"
	ItemTemplate = ItemTemplate & "</div>"
	ItemTemplate = ItemTemplate & "<div class=""rre_dots""><img src=""" & templatepath & "images/dots.png"" width=""238"" height=""1"" border=""0"" alt=""""/></div>"
	' ================================================
	
	On Error Resume Next
	Set xmlHttp = Server.CreateObject("MSXML2.XMLHTTP.3.0")
	xmlHttp.Open "Get", URLToRSS, false
	IF Err.number <> 0 then
		Response.Write "The URL is not valid"
		Exit Sub            
	END IF        
	xmlHttp.Send()
	RSSXML = xmlHttp.ResponseText
	
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = false
	xmlDOM.LoadXml(RSSXML)
	IF Err.number <> 0 then
		Response.Write "The RSS Data is not valid"
		Exit Sub            
	END IF        

	Set RSSItems = xmlDOM.getElementsByTagName("item") ' collect all "items" from downloaded RSS
	'Response.Write "<table>"
	j = -1
	FOR EACH RSSItem IN RSSItems
		FOR EACH child IN RSSItem.childNodes
			Select case lcase(child.nodeName)
			 case "title"
				   RSStitle = child.text
			 case "link"
					'if child.getAttribute("rel") = "alternate" then
					'	RSSlink = child.getAttribute("href")
					'end if
			 case "feedburner:origlink"
					RSSlink = child.text
			 case "description"
				   RSSdescription = child.text
				   
				   'Set re = New RegExp
				   're.Pattern = "<[^>]+>"
				   're.IgnoreCase = true
				   're.Global = true
				   
				   'RSSdescription = re.Replace(RSSdescription, "")
				   
				   'set re = nothing
					
					'if len(RSSdescription) > 200 then RSSdescription = mid(RSSdescription, 1, 200) & "..."
			 case "pubdate"
					mydate = split(child.text, " ")
					
					rssdate = mydate(0) & " " & mydate(1) & " " & mydate(2) & " " & mydate(3)
					
					'mydate = cdate(child.text)
					'rssDate = mydate & ""
				   'RSSdate = split(RSSDate, "T")(0)
				   'mon = Month(RSSdate)
				   'd = Day(RSSdate)
				   'yr = Year(RSSdate)
				   
				   'RSSdate = GetMonth(mon) & " " & d & ", " & yr
			case "dc:creator"
					RSScreator = replace(child.text, "_", " ")
			END SELECT
		NEXT                                     
		j = J+1                                     
		IF J < CInt(MaxNumberOfItems) THEN
			Itemcontent = ItemTemplate
			Itemcontent = Replace(Itemcontent,"{DATE}", RSSdate)
			ItemContent = Replace(ItemContent,"{LINK}",RSSlink)
			ItemContent = Replace(ItemContent,"{TITLE}",RSSTitle)
			ItemContent = Replace(ItemContent,"{DESCRIPTION}",RSSDescription)
			ItemContent = Replace(ItemContent,"{CREATOR}",RSSCreator)
			
			Response.Write ItemContent
			ItemContent = ""
		END IF
	NEXT
	'Response.Write "</table>"
	On Error GoTo 0
END SUB
%>		
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		
		<% 
			MetaTagSub
			MenuSubJS
			%>		
		<LINK REL="SHORTCUT ICON" HREF="<% = TemplatePath %>favicon.ico">
		<link href="<% = TemplatePath %>styles.css" rel="stylesheet" type="text/css" />
		<link href="<% = TemplatePath %>menu.css" rel="stylesheet" type="text/css" />
		<!--[if lt IE 7]>			
			<link href="<% = TemplatePath %>styles_ie6.css" rel="stylesheet" type="text/css" />
		<![endif]-->

		<script type="text/javascript" src="ASPModules/searchboxes4.js"></script>	
		<script type="text/javascript" src="<%=templatepath%>slidebox.js"></script>
		<script type="text/javascript" src="/aspmodules/lw_afp_slide.js"></script>
		<script src="<% = TemplatePath %>jquery.ztwitterfeed.min.js" type="text/javascript"></script>		

		<script type="text/javascript" src="<% = TemplatePath %>fancybox/jquery.mousewheel-3.0.2.pack.js"></script>
		<script type="text/javascript" src="<% = TemplatePath %>fancybox/jquery.fancybox-1.3.1.js"></script>
		<link rel="stylesheet" type="text/css" href="<% = TemplatePath %>fancybox/jquery.fancybox-1.3.1.css" media="screen" />
		<link rel="stylesheet" type="text/css" href="<% = TemplatePath %>stylefb.css" media="screen" />
		<script language="javascript">
			var mySlider; // Rotate - slide	
			$().ready(function() { // Any Rotate/Carousel
				mySlider = new LW_Slider("slider",-1,"left",600); // Rotate - slide	
				
				$("#various1").fancybox({
					'titlePosition'		: 'inside',
					'transitionIn'		: 'swing',
					'width'             : '600',
					'height'             : '400',
					'modal'				: true,
					'transitionOut'		: 'none'
				});
				$("#rre_follow_us").slideBox({
					width: "100%", 
					height: "334px", 
					position: "top"
				});
				
				$('#rre_tweets').twitterfeed('HomesinBG', {
					limit: 1
				});
				
				$("#togglelink").mouseover(function(){
					$("#togglethis").show("slow");
				});
				
				$("#togglelink_close").mouseout(function(){
					$("#togglethis").hide("slow");
				});
				
			}); // Any Rotate
		</script>
		<script type="text/javascript" src="<% = TemplatePath %>thickbox.js"></script>
		<link rel="stylesheet" href="<% = TemplatePath %>thickbox.css" type="text/css" media="screen" />
	</head>
	<body>
		<div class="rre_wrapper">
			<div class="rre_branding">
				<div id="rre_follow_us" style="display:none;">
					<div id="rre_follow_us_contact">
					
						<div id="rre_follow_us_social_title">Contact Us</div>
		
					<a href="contactus.asp"><img src="<%=templatepath%>images/office_img.jpg" border="0" width="250" height="145" alt="Click here to view office on map"/></a>
						<div class="clear_10"></div>	
						<div>
						<% = Officename%><br/>
						<% = Address%> <br/>
						 <% =City  %>,<% =State  %> 
						 <br/>
						 <% =Zip  %>
						</div>
						
						<div class="rre_follow_us_spacer">
							Phone: <% = Officephone%>
						</div>
						<div class="rre_follow_us_spacer">
							<a href="contactus.asp">Contact Us Email Form &gt;&gt;</a>
						</div>
					</div>
					<div id="rre_follow_us_icons">
						<div id="rre_follow_us_social_title">Social Media Links</div>
							<a href="http://www.facebook.com/pages/REMAX-Real-Estate-Executives/103963969652484?ref=sgm" target="_blank" class="rre_follow_us_facebook">Friend Us On Facebook</a>
							<a href="http://www.twitter.com/HomesinBG" target="_blank" class="rre_follow_us_twitter">Follow Us On Twitter</a>
							<a href="http://www.youtube.com/user/HomesInBowlingGreen" target="_blank" class="rre_follow_us_youtube">Watch Our Youtube Channel</a>
							<a href="http://www.linkedin.com/pub/homesinbowlinggreen-re-max-real-estate-executives/1a/b7/109" target="_blank" class="rre_follow_us_linkedin">View Our LinkedIn Profile</a>
							<a href="http://bowlinggreenmarketreport.com/blog" target="_blank" class="rre_follow_us_blog">Visit The RE/MAX Real Estate Executives Market Report</a>
					</div>
						<div id="rre_follow_us_twitter">
						<div class="clear_150"></div>
						<div id="twitter_wrap">
							<div id="rre_tweets"></div><!--tweet will populate here-->
						</div>
					</div>
				</div>
				<a class="rre_logo"  href="default.asp"><img src="<%=templatepath%>images/spacer.gif" width="250" height="141" alt="click here to return to homepage" border="0"/></a>

						<div class="rre_agentHeader">
				<div style="background:url(<%=PhotoSrc%>) no-repeat center;" class="rre_agentPhoto"/></div>
			      <div class="rre_agentInfo">
				    <span class="rre_agentName">
				      <a href="/view_agent.asp"><% = AgentName %></a> 
				    </span>
				    <span class="rre_justRewards">
				      <% = JustAwards%>
				    </span>
				    <span class="rre_agentEmail">
				      <a href="mailto:<% = Email %>"><% = Email %></a>
				    </span>
					<span class="rre_agentPhone">
					<% IF directline = "" then%>
						<% IF officephone <> "" THEN %>
						Phone:&nbsp;<% = officephone %>
						<% IF Extension <> "" THEN %>
						Ext&nbsp;<% = Extension %>
					<% END IF %> 
					<% END IF %> 
					<% ELSE %>
						Phone:&nbsp;<% = directline %>
					<% END IF %>
					<% IF mobile = "" THEN %>
					 <% ELSE %>
						<br/>Mobile:&nbsp;<% = mobile %>
					<% END IF %> 
					<% IF agentfax = "" THEN %>
					<% ELSE %>
						<br/>Fax:&nbsp;<% =agentfax %>
					<% END IF %> 
					<% IF Tollfree = "" THEN %>
					<% ELSE %>
						<br/>Toll Free:&nbsp;<% = TollFree %>
					<% END IF %> 
					</span><!--agentPhone-->
				  </div><!-- rre_agentInfo-->
				</div><!--rre_agentHeader-->
				
				
				
			</div>
		
			<!--end of the branding-->
			<div class="rre_navigation">
				<%MenuWebLocation = "topnav" %>	
				<% MenuSub %>
				
			</div>
			<!--end of the navigation-->
			
			<div class="clear"></div>
			<div class="rre_hp_container">
				<div class="rre_hp_search_box">
				<span class="rre_qs_title">Find Your New Dream Home</span>
					<div class="rre_hp_search">
						<form name="frm" action="viewhome.asp" id="frm">
						<input name="HOMEPAGESEARCH" id="HOMEPAGESEARCH" value="TRUE" type="hidden"/>
						<input name="NoSummary" value="Y" type="hidden"/>
						<input name="Search" value="" type="hidden"/>
						<input name="SearchType" value="" type="hidden"/>
						<input name="HomeStatusFilter" value="'1','4'" type="hidden"/>
						<input type="hidden" name="MapMyResults" value="" />
						<%
						
						QS_SearchDefID = QS_GetSearchDefID()
						QS_FirstTabIndex = 1
						maxquicksearchfields = 6												
						%>	
						<table class="rre_qsearch_table" width="365" border="0" cellpadding="2" cellspacing="3">
							<tr>
								<td colspan="3">
									<span class="rre_location_text">Location</span>
									<%
									QS_GetCityStateZip_FieldName = "pCSLocation"
									QS_GetCityStateZip_NameFieldName = "tmpCity"
									QS_GetCityStateZip_FieldClassName = "rre_location"
									QS_GetCityStateZip_SearchCity = "Y"
									QS_GetCityStateZip_SearchCounty = "Y"
									QS_GetCityStateZip_SearchState = "Y"
									QS_GetCityStateZip_SearchZip = "Y"
									success = QS_GetCityStateZip()								        						        
									%>
									</td>
									<td>
									<span class="rre_proptype_text">Property Type</span>
									<%
										QS_GetPropertyType_ControlClassName = "rre_proptype" 
										success = QS_GetPropertyType()
									%>
								</td>
							</tr>
							<tr class="row_padding"><td colspan="4"></td></tr>
							<tr>
								<td>
									<span class="rre_proptype_text">Beds</span>
									<select id="totbed"  name="totbed" class="rre_beds">
										<option value="0" selected="selected"> Any </option>
										<option value="1">1</option>
										<option value="2">2</option>
										<option value="3">3</option>
										<option value="4">4</option>
										<option value="5">5</option>
										<option value="6">6</option>
										<option value="7">7</option>
										<option value="8">8</option>
										<option value="9">9</option>
										<option value="10">10</option>
									</select>
								</td>
								<td>
								<span class="rre_proptype_text">Baths</span>
									<select id="totbath"  name="totbath" class="rre_baths">
										<option value="0" selected="selected"> Any </option>
										<option value="1">1</option>
										<option value="2">2</option>
										<option value="3">3</option>
										<option value="4">4</option>
										<option value="5">5</option>
										<option value="6">6</option>
										<option value="7">7</option>
										<option value="8">8</option>
										<option value="9">9</option>
										<option value="10">10</option>
									</select>
								</td>
								<td>
								<span class="rre_proptype_text">Min Price</span>
									<select class="rre_min_price" name="minprice" id="minprice">
										<option value="0" selected="selected">No Minimum </option>
										<option value="50000">$50,000</option>
										<option value="60000">$60,000</option>
										<option value="70000">$70,000</option>
										<option value="80000">$80,000</option>
										<option value="90000">$90,000</option>
										<option value="100000">$100,000</option>
										<option value="125000">$125,000</option>
										<option value="150000">$150,000</option>
										<option value="175000">$175,000</option>
										<option value="200000">$200,000</option>
										<option value="225000">$225,000</option>
										<option value="250000">$250,000</option>
										<option value="275000">$275,000</option>
										<option value="300000">$300,000</option>
										<option value="325000">$325,000</option>
										<option value="350000">$350,000</option>
										<option value="375000">$375,000</option>
										<option value="425000">$425,000</option>
										<option value="400000">$400,000</option>
										<option value="450000">$450,000</option>
										<option value="475000">$475,000</option>
										<option value="500000">$500,000</option>
										<option value="550000">$550,000</option>
										<option value="600000">$600,000</option>
										<option value="650000">$650,000</option>
										<option value="700000">$700,000</option>
										<option value="750000">$750,000</option>
										<option value="800000">$800,000</option>
										<option value="850000">$850,000</option> 
										<option value="900000">$900,000</option>
										<option value="950000">$950,000</option>
										<option value="1000000">$1,000,000</option>
										<option value="1100000">$1,100,000</option>
										<option value="1200000">$1,200,000</option>
										<option value="1300000">$1,300,000</option>
										<option value="1400000">$1,400,000</option>
										<option value="1500000">$1,500,000</option>
										<option value="1600000">$1,600,000</option>
										<option value="1700000">$1,700,000</option>
										<option value="1800000">$1,800,000</option>
										<option value="1900000">$1,900,000</option>
										<option value="2000000">$2,000,000</option>
										<option value="1100000">$1,100,000</option>
										<option value="1200000">$1,200,000</option>
										<option value="1300000">$1,300,000</option>
										<option value="1400000">$1,400,000</option>
										<option value="1500000">$1,500,000</option>
										<option value="1600000">$1,600,000</option>
										<option value="1700000">$1,700,000</option>
										<option value="1800000">$1,800,000</option>
										<option value="1900000">$1,900,000</option>
										<option value="2000000">$2,000,000</option>
										<option value="2500000">$2,500,000</option>
										<option value="3000000">$3,000,000</option>
										<option value="3500000">$3,500,000</option>
										<option value="4000000">$4,000,000</option>
										<option value="4500000">$4,500,000</option>
										<option value="5000000">$5,000,000</option>
										<option value="6000000">$6,000,000</option>
										<option value="7000000">$7,000,000</option>
										<option value="8000000">$8,000,000</option>
										<option value="9000000">$9,000,000</option>
										<option value="10000000">$10,000,000</option>
									</select>
								</td>
								<td><span class="rre_proptype_text">Max price</span>
								<select class="rre_max_price" name="maxprice" id="maxprice">
									<option value="2100000000" selected="selected">No Maximun</option>
									<option value="50000">$40,000</option>
									<option value="60000">$60,000</option>
									<option value="80000">$80,000</option>
									<option value="100000">$100,000</option>
									<option value="125000">$125,000</option>
									<option value="150000">$150,000</option>
									<option value="175000">$175,000</option>
									<option value="200000">$200,000</option>
									<option value="225000">$225,000</option>
									<option value="250000">$250,000</option>
									<option value="275000">$275,000</option>
									<option value="300000">$300,000</option>
									<option value="325000">$325,000</option>
									<option value="350000">$350,000</option>
									<option value="375000">$375,000</option>
									<option value="400000">$400,000</option>
									<option value="425000">$425,000</option>
									<option value="450000">$450,000</option>
									<option value="475000">$475,000</option>
									<option value="500000">$500,000</option>
									<option value="550000">$550,000</option>
									<option value="600000">$600,000</option>
									<option value="650000">$650,000</option>
									<option value="700000">$700,000</option>
									<option value="750000">$750,000</option>
									<option value="800000">$800,000</option>
									<option value="850000">$850,000</option> 
									<option value="900000">$900,000</option>
									<option value="950000">$950,000</option>
									<option value="1000000">$1,000,000</option>
									<option value="1100000">$1,100,000</option>
									<option value="1200000">$1,200,000</option>
									<option value="1300000">$1,300,000</option>
									<option value="1400000">$1,400,000</option>
									<option value="1500000">$1,500,000</option>
									<option value="1600000">$1,600,000</option>
									<option value="1700000">$1,700,000</option>
									<option value="1800000">$1,800,000</option>
									<option value="1900000">$1,900,000</option>
									<option value="2000000">$2,000,000</option>
									<option value="2500000">$2,500,000</option>
									<option value="3000000">$3,000,000</option>
									<option value="3500000">$3,500,000</option>
									<option value="4000000">$4,000,000</option>
									<option value="4500000">$4,500,000</option>
									<option value="5000000">$5,000,000</option>
									<option value="6000000">$6,000,000</option>
									<option value="7000000">$7,000,000</option>
									<option value="8000000">$8,000,000</option>
									<option value="9000000">$9,000,000</option>
									<option value="10000000">$10,000,000</option>
								</select>
								</td>
							</tr>
							<tr>
								<td class="rre_daysold" align="right" colspan="4">Within the last:
								<select class="rre_beds" name="daysold" id="daysold">
									<option value="1">1</option>
									<option value="2">2</option>
									<option value="3">3</option>
									<option value="4">4</option>
									<option value="5">5</option>
									<option value="6">6</option>
									<option value="7">7</option>
									<option value="14">14</option>
									<option value="21">21</option>
									<option value="30">30</option>
									
								</select>
								days.
							</td>
								
							</tr>
						</table>
						<div class="rre_qsearch_btns">
							<span class="rre_btn_text">View Your Results By:</span>
							<a class="rre_list_btn" href="javascript:SearchClick();" id="map" ><img src="<% = TemplatePath %>images/spacer.gif" height="32"  width="79" border="0" alt="Click here to view results by list"/></a>
							<a class="rre_map_btn" href="javascript:MapSearchClick();"><img src="<% = TemplatePath %>images/spacer.gif" height="32"  width="81" border="0" alt="Click here to view results on the map"/></a>
						</div>
					</form>
				</div>
					<div class="rre_featured_property_box">
					<% 'Rotate - Slide
						MLSPhotoWidth = "260"
						FilterStatus = "1,3,4" 
						FeaturedPropertyWindow "slider",10,false
					%>
						<div class="rre_slider">
							<a class="slider_controls_prev" href="#" onclick="mySlider.prevslide(); return false;"><img src="<%=templatepath%>images/prev.jpg" width="64" height="30" border="0" alt="click here to view next listing"/></a>
							<a class="slider_controls_next" href="#" onclick="mySlider.nextslide(); return false;"><img src="<%=templatepath%>images/next.jpg" width="64" height="30" border="0" alt="click here to view next listing"/></a>
						</div>
					</div>
				</div>
				<!--end of top search-->
			</div>
			<!--end of the container-->
			<div class="clear_5"></div>
			<div class="rre_middle_row">
				<div class="rre_mid_box">
					<div class="rre_home_finder">
						<div class="rre_middle_title"><a href="login.asp">Home Finder</a></div>
						<p>Tell us what you are looking for and we'll email you homes that meet your criteria.</p>
						<a class="mid_btn" href="login.asp">Login / Sign Up</a>
					</div>
				</div>
				<div class="rre_mid_box">
					<div class="rre_schools">
						<div class="rre_middle_title"><a href="infotopic.asp?InfoID=20">School Info</a></div>
						<p>Get the information you need about schools in our area.</p>
						<a class="mid_btn" href="infotopic.asp?InfoID=20">View Schools</a>
					</div>
				</div>
				<div class="rre_mid_box">
					<div class="rre_new_home">
						<div class="rre_middle_title"><a href="infotopic.asp?InfoID=31">New Home Communities</a></div>
						<p>Get the knowledge you need when beginning the endeavor of building a new home.</p>
						<a class="mid_btn2" href="infotopic.asp?InfoID=31">New Home Communities</a>
					</div>
				</div>
				<div class="rre_mid_box2">
					<div class="rre_mobile">
								<div class="rre_mobile_title"><a href="<% = TemplatePath %>MobileWolf.htm?KeepThis=true&TB_iframe=true&height=400&width=604"><img src="<%=templatepath%>images/spacer.gif" alt="click here to learn more about mobile wolf" width="130" height="30" border="0"/></a></div>
								<div class="mobile_text">Text: <span class="text_info"> 32323</span>and type in <span class="text_info"> <i>moreinfo</i></span> &nbsp; �a space� and the house # to view more info on all homes in the Bowling Green area.
								</div>
								<a class="btm_btn thickbox" href="<% = TemplatePath %>MobileWolf.htm?KeepThis=true&TB_iframe=true&height=400&width=604">Search Now</a>
							</div>						
				</div>
			</div>
			<div class="clear_10"></div>
			<div class="rre_hp_btm">
				<div class="rre_hp_left">
					<div class="rre_hp_left_wrap">
						<div class="rre_agent_content">
							<% = info %>
						</div>	
						<!--end of agent content-->
						<div class="rre_hp_left_bottom"></div>
					</div><!--end of the top_left"-->
				</div><!--end of the left -->
			</div>
			<!--end of the rre_hp btm-->
		</div>
		<!--end of wrapper-->
		
		<div class="rre_footer_wrapper">
			<div class="rre_footer">
				<a class="rre_agent" href="http://<%=ExtranetProjectUrl%>">Agent Login</a>
				<span class="rre_sign">&nbsp;</span>
				<div class="rre_btm_nav">
					<%MenuWebLocation = "bottomnav" %>	
					<% MenuSub %>
				</div>
			</div>
			<div class="rre_disclaimer">
			<%disclaimersub%>
			
			</div>
		</div>
		<!--end of the footer wrapper-->
	</body>
</html>
<%
function GetMonth(m)
	dim retval
	select case m & ""
		case "1"
			retval = "January"
		case "2"
			retval = "February"
		case "3"
			retval = "March"
		case "4"
			retval = "April"
		case "5"
			retval = "May"
		case "6"
			retval = "June"
		case "7"
			retval = "July"
		case "8"
			retval = "August"
		case "9"
			retval = "September"
		case "10"
			retval = "October"
		case "11"
			retval = "November"
		case "12"
			retval = "December"
		
	end select
	GetMonth = retval
end function
%>
