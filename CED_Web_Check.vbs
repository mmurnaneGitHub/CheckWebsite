'========================================================================== 
' 
' NAME: GADS Web Page Monitor 
' 
' AUTHOR: Mike Murnane 
' DATE  : 5/3/2019
' 
' DESCRIPTION: Monitor GADS web pages and sends email of results.
' Path: \\Geobase-win\CED\GADS\R2017\R426\ScheduledTask\CED_Web_Check.vbs
' 
' Pages to monitor - Deep link will not work if all elements of page haven't loaded (geocoder - need to do in steps)
  ' Getting 401 Unauthorized (The request requires user authentication.) for the following:
               '"http://geobase-dbnewer/website/labels/index.html", _
               '"http://tpd-as001.tacoma.lcl/website/Police/CrackTrack/edit.htm", _
' ERROR: msxml6.dll: The server returned an invalid or unrecognized response | Popups with errors on the page will cause this error.'
'        Example: Using https instead of http on staff DART map will have a popup that not all layers could be added - 
'                 http://tacoma.maps.arcgis.com/home/webmap/viewer.html?webmap=6971a49654b7419f916343c202e61827'
'                 https://www.prairielinetrail.org/map/
'Following links (ES & Smoke Test) throwing 401 errors:
'               "http://dbmsql50/ReportServer/Pages/ReportViewer.aspx?/EnvironmentalServices/Reports/Operations_Maintenance/Transmission/GIS_MAIN_INFO&rs:Command=Render&rc:Parameters=false&FACILITYID=6251593", _
'               "http://wsgovme02/Projects/smoketest.asp?Addressid=1738", _
'========================================================================== 
Dim count, badURL, goodMsg
count = 0
badURL = ""
goodMsg = "SUCCESS! All pages successfully returned a HTTP status of 200."

theUrls = Array("http://wsitd03/website/DART/StaffMap/", _
               "https://www.cityoftacoma.org/maps", _ 
               "https://wspdsmap.cityoftacoma.org/website/Art2/viewer.htm", _ 
               "https://wspdsmap.cityoftacoma.org/website/ArtAtWork/StudioMap.htm", _
               "https://wspdsmap.cityoftacoma.org/website/BLUS/StreetView.htm", _ 
               "https://wspdsmap.cityoftacoma.org/website/CED/LEAP/", _ 
               "https://wspdsmap.cityoftacoma.org/website/CMO/TacomaCouncil.htm", _ 
               "https://wspdsmap.cityoftacoma.org/website/CMO2/PaidLeave/", _ 
               "https://wspdsmap.cityoftacoma.org/website/CMO2/Top10/MapTour/", _ 
               "https://wspdsmap.cityoftacoma.org/website/DART/staff/TacomaPermitsMap.htm", _ 
               "https://wspdsmap.cityoftacoma.org/website/Downtown/Downtown.htm", _ 
               "https://wspdsmap.cityoftacoma.org/website/Finance/Verify/", _
               "https://wspdsmap.cityoftacoma.org/website/FLUM/", _ 
               "https://wspdsmap.cityoftacoma.org/website/Google/Drag_n_Drop/", _ 
               "https://wspdsmap.cityoftacoma.org/website/Google/StreetView/", _
               "https://wspdsmap.cityoftacoma.org/website/GreenMap/", _ 
               "https://wspdsmap.cityoftacoma.org/website/HistoricIs/", _ 
               "https://wspdsmap.cityoftacoma.org/website/HistoricMap/", _ 
               "https://wspdsmap.cityoftacoma.org/website/NCS/Cleanup/", _
               "https://wspdsmap.cityoftacoma.org/website/PDS/BuildingTacoma/", _ 
               "https://wspdsmap.cityoftacoma.org/website/PDS/LandUse/", _ 
               "https://wspdsmap.cityoftacoma.org/website/PDS/MJ/", _  
               "https://wspdsmap.cityoftacoma.org/website/PDS/OneTacoma/", _ 
               "https://wspdsmap.cityoftacoma.org/website/PDS/Permits/", _
               "https://wspdsmap.cityoftacoma.org/website/PDS/PermitDashboard/", _
               "https://wspdsmap.cityoftacoma.org/website/PDS/LandUse/", _
               "https://wspdsmap.cityoftacoma.org/website/PDS/Tideflats/", _
               "https://wspdsmap.cityoftacoma.org/website/PDS/TIR/", _
               "https://wspdsmap.cityoftacoma.org/website/PDS/Zoning/", _ 
               "https://wspdsmap.cityoftacoma.org/website/PoetryTour/", _ 
               "https://wspdsmap.cityoftacoma.org/website/PW/ParkingCalculator2/", _
               "https://wspdsmap.cityoftacoma.org/website/PW/RPP/", _ 
               "https://wspdsmap.cityoftacoma.org/website/spaceworks/viewer.htm", _ 
               "https://wspdsmap.cityoftacoma.org/website/tacomaspace/", _ 
               "https://www.cityoftacoma.org/cms/One.aspx?portalId=169&pageId=123158", _ 
               "https://www.bing.com/maps/?v=2&cp=47.255864291960876~-122.44179027140726&lvl=19&sty=b", _
               "https://wspdsmap.cityoftacoma.org/website/BLUS/StreetView.htm", _
               "http://cms.cityoftacoma.org/Planning/Zoning%20Reference%20Guide%202016.pdf", _
               "http://cms.cityoftacoma.org/cityclerk/Files/MunicipalCode/Title13-LandUseRegulatoryCode.PDF", _
               "https://wspdsmap.cityoftacoma.org/website/HistoricMap/scripts/summary.asp?ID=(490)&map=(47.258287,-122.44671)", _
               "https://wsowa.ci.tacoma.wa.us/cot-itd/addressbased/permithistory.aspx?Address=747%20MARKET%20ST&Mode=simple", _
               "https://epip.co.pierce.wa.us/CFApps/atr/epip/summary.cfm?parcel=9005250030", _
               "http://www.govme.org/Common/MyTacoma/MyTacoma.aspx?Parcel=9005250030", _
               "http://www.govme.org/gMap/Info/eVaultFilter.aspx?StreetIDs=10751", _
               "https://fortress.wa.gov/ecy/dirtalert/", _
               "https://tacoma.maps.arcgis.com/home/webmap/viewer.html?webmap=04a559793bdf43998c4e8f5f3b8a4e4d", _
               "https://tacoma.maps.arcgis.com/home/webmap/viewer.html?webmap=35fde3a26d8b47f288cf7f81a915c09c", _
               "https://tacoma.maps.arcgis.com/home/webmap/viewer.html?webmap=3a1add3b4de947f2bf1f19b83a8c2266", _
               "https://tacoma.maps.arcgis.com/home/webmap/viewer.html?webmap=4bf2dfbb4aa642cbad278f3192387612", _
               "http://tacoma.maps.arcgis.com/home/webmap/viewer.html?webmap=6971a49654b7419f916343c202e61827", _
               "https://tacoma.maps.arcgis.com/home/webmap/viewer.html?webmap=6e03e8c26aad4b9c92a87c1063ddb0e3", _
               "https://tacoma.maps.arcgis.com/home/webmap/viewer.html?webmap=6e03e8c26aad4b9c92a87c1063ddb0e3", _
               "https://tacoma.maps.arcgis.com/home/webmap/viewer.html?webmap=73ea9b33b89f42128c99628c6bf49d9e", _
               "https://tacoma.maps.arcgis.com/home/webmap/viewer.html?webmap=76f1a118a00f4587bd41d44ed6cb3950", _
               "https://tacoma.maps.arcgis.com/home/webmap/viewer.html?webmap=7865507a1def49638bd635defece9378", _
               "https://tacoma.maps.arcgis.com/home/webmap/viewer.html?webmap=a0e087e7b7844e27a4db793ac8c32ce5", _
               "https://tacoma.maps.arcgis.com/home/webmap/viewer.html?webmap=c616cb270a634fa48d51e3333fdbc67d", _
               "https://tacoma.maps.arcgis.com/home/webmap/viewer.html?webmap=ca96fdd89a4d4a19a2a463fefaec7a6b", _
               "https://tacoma.maps.arcgis.com/home/webmap/viewer.html?webmap=e547ab1a58804d19acd9a0059fb6a5ae") 
                
' Loop through and test pages  
For Each Url In theUrls 
    GetWebPage Url
    count = count + 1 
Next 
 
Sub GetWebPage(ByVal Url) 
    ' Use latest version version  
    Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")   'The standard request fired from a local machine forbids access to sites that aren't trusted by IE. This is the Server side object that doesn't perform those checks.
    oXMLHTTP.open "GET", Url 
    oXMLHTTP.send 
    'msgbox "Response: " & oXMLHTTP.status
    ' Test for response code 200 
    If Not oXMLHTTP.status = 200 Then 
        'SendEmail Url, oXMLHTTP.status
        badURL = badURL & vbCrLf & Url& " - Status: " & oXMLHTTP.status
    End If 
End Sub 

SendEmail(badURL) 
 
'Sub SendEmail(Msg, status) 
Sub SendEmail(badURL) 
    ' Set SMTP server 
    EmailSrv = "smtp001.tacoma.lcl" 
    EmailFrom = "mmurnane@cityoftacoma.org" 
    EmailTo = "mmurnane@cityoftacoma.org" 
    Set objEmail = CreateObject("CDO.Message") 
    objEmail.From = EmailFrom 
    objEmail.To = EmailTo 
    objEmail.Subject = "GADS Web Pages Monitor" 
    
    'Check if any bad pages found
    If Len(badURL)>0 Then
      theMsg = "PROBLEM FOUND! The following GADS web pages are not responding: " & badURL
    Else
      theMsg = goodMsg    
    End If   

    objEmail.Textbody = count & " GADS web pages checked." & vbCrLf & vbCrLf _
      & theMsg & vbCrLf & vbCrLf _
      & "HTTP Status Code Definitions: https://www.w3.org/Protocols/rfc2616/rfc2616-sec10.html" & vbCrLf & vbCrLf _
      & "Log Files: \\Geobase-win\CED\GADS\R2017\R426\ScheduledTask\log" & vbCrLf & vbCrLf _
      & "Rerun this report: \\Geobase-win\CED\GADS\R2017\R426\ScheduledTask\CED_Web_Check.bat "
      
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = EmailSrv  
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
    objEmail.Configuration.Fields.Update 
    objEmail.Send 
End Sub 

'Clear the objects
Set oXMLHTTP = Nothing
Set objEmail = Nothing