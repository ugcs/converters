Option Explicit

if 2 <> WScript.Arguments.Count then
	MsgBox "Please specify KML file name and route template in command line."
	WScript.Quit
end if

kmlFile2Routes WScript.Arguments.Item(0), WScript.Arguments.Item(1)

WScript.Quit

'----------------------------------------------------------------------------------
' Converts KML file to UgCS route files (each Placemark to separate UgCS route file)

Sub kmlFile2Routes( kmlFileName, routeTemplateFileName)
	
	Dim xmlDoc
	Dim placemarks
	Dim i

	WScript.Echo "KML file name: " & kmlFileName
	WScript.Echo "Template file name: " & routeTemplateFileName

	Set xmlDoc = CreateObject( "Msxml2.DOMDocument" )
	xmlDoc.Async = false

	if false = xmlDoc.load(kmlFileName) then
		MsgBox "Can't load KML file: " & xmlDoc.parseError.reason
		exit sub
	end if

	xmlDoc.setProperty "SelectionLanguage", "XPath"
	xmlDoc.setProperty "SelectionNamespaces", "xmlns:kml='http://www.opengis.net/kml/2.2'"

	Set placemarks = xmlDoc.selectNodes("//kml:Placemark")
	if 0 = placemarks.length then
		MsgBox "Can't find any Placemark in KML file!"
		exit sub
	end if

	for i = 0 to placemarks.length - 1
		placemark2route placemarks(i), i, routeTemplateFileName
	next

End Sub

'----------------------------------------------------------------------------------
' Converts one KML Placemark to UgCS route

Sub placemark2route( kmlPlacemark, placemarkIndex, routeTemplateFileName)

	Dim kmlPolygonName
	Dim kmlCoordinates

	Dim routeName
	Dim coordsArray
	Dim longtitude
	Dim lattitude
	Dim altitude
	Dim i

	Dim ugcsRouteTemplate
	Dim ugcsRouteFileName	

	' Try to get KML placemark name - it will be used as UgCS route and file name
	Set kmlPolygonName = kmlPlacemark.selectSingleNode("kml:name")
	if kmlPolygonName is Nothing then
		routeName = getFileNameFromUrl( kmlPlacemark.ownerDocument.url) & "." & placemarkIndex
		WScript.Echo "Found unnamed placemark"
	else
		routeName = kmlPolygonName.Text
		WScript.Echo "Found placemark '" & routeName & "'"
	end if

	Set ugcsRouteTemplate = getUgcsRouteTemplate( routeTemplateFileName)
	if ugcsRouteTemplate is Nothing then
		exit sub
	end if

	' Update route name in route template
	ugcsRouteTemplate.selectSingleNode("//Route/name").setAttribute "v", routeName


	' Try to find coordinates node inside of placemark
	Set kmlCoordinates = kmlPlacemark.selectSingleNode("//kml:coordinates")
	if kmlCoordinates is Nothing then
		WScript.Echo " this placemark does not contain any coordinates, exiting..."
		exit sub
	end if

	coordsArray = Split( kmlCoordinates.Text, " ")

	for i = 0 to UBound(coordsArray)

		if false = parseCoordinate( coordsArray(i), longtitude, lattitude, altitude) then
			MsgBox "Can't parse coordinates string '" & coordsArray(i) & "' in placemark '" & routeName & "'!"
		else
			addUgcsFigurePoint ugcsRouteTemplate, longtitude, lattitude
		end if
	next

	' Save modified UgCS route template as new route 
	ugcsRouteFileName = routeName & ".xml"
	ugcsRouteTemplate.save ugcsRouteFileName
	WScript.Echo " UgCS route saved in file '" & ugcsRouteFileName & "'"

End Sub

'----------------------------------------------------------------------------------
' Loads UgCS route template from file and clears collection of figure 
' (AreaScan/Perimeter/Photogrammetry tool) points
' Returns:
'	success: route template (Msxml2.DOMDocument)
'	error: Nothing

function getUgcsRouteTemplate( routeTemplateFileName)

	Dim xmlDoc
	Dim points

	Set getUgcsRouteTemplate = Nothing

	Set xmlDoc = CreateObject( "Msxml2.DOMDocument" )
	xmlDoc.Async = false
	xmlDoc.preserveWhiteSpace = true

	if false = xmlDoc.load( routeTemplateFileName) then
		MsgBox "Can't load UgCS route template file '" & routeTemplateFileName & "': " & xmlDoc.parseError.reason
		exit function
	end if

	xmlDoc.setProperty "SelectionLanguage", "XPath"

	' find first figure definition 
	Set points = getUgcsFigurePoints( xmlDoc)
	if points is Nothing then
		MsgBox "Can't find figure definition in template file '" & routeTemplateFileName & "'!"
		exit function
	end if
	
	' clear all points from polygon definition
	points.selectNodes("o").removeAll()

	Set getUgcsRouteTemplate = xmlDoc 

end function

'----------------------------------------------------------------------------------
' Returns figure points list node from UgCS route

function getUgcsFigurePoints( ugcsRouteDomDocument)

	' find first figure definition 
	Set getUgcsFigurePoints = ugcsRouteDomDocument.selectSingleNode("//Route/segments/figure/ugcs-List")

end function

'----------------------------------------------------------------------------------
' Adds new point to UgCS route figure definition

sub addUgcsFigurePoint( ugcsRouteDomDocument, longtitude, lattitude)

	const PI = 3.14159265358979 

	Dim points
	Dim point

	' Convert longtitude and lattitude to radians
	longtitude = longtitude * PI / 180
	lattitude = lattitude * PI / 180

	Set points = getUgcsFigurePoints( ugcsRouteDomDocument)

	' create and fill new point
	' it should have format 
	' <o v7="AGL" v6="0.0" v4="longtitude" v3="lattitude" v2="order"/>

	Set point = ugcsRouteDomDocument.createElement("o")
	point.setAttribute "v7", "AGL"
	point.setAttribute "v6", "0.0"
	point.setAttribute "v4", longtitude
	point.setAttribute "v3", lattitude
	point.setAttribute "v2", points.selectNodes("o").length

	points.appendChild point

	' trick - UgCS XML parser can't load XML with more then one element per line
	' so will add text node with new line
	points.appendChild ugcsRouteDomDocument.createTextNode( VbCrLf)

end sub

'----------------------------------------------------------------------------------
' Parses single coordinate string in format "longtitude,lattitude,altitude"
' Return:
'	true - parsing was succesful
'	false - some error

function parseCoordinate( coordString, lonOut, latOut, altOut)
	
	Dim coordinateComponents

	parseCoordinate = false
	
	coordinateComponents = Split( coordString, ",")

	' check if have at least longtitude and lattitude here
	if UBound(coordinateComponents) < 1 then
		exit function
	end if

	lonOut = str2double( coordinateComponents(0))
	latOut = str2double( coordinateComponents(1))

	if 2 = UBound(coordinateComponents) then	' coordinates with altitude
		altOut = str2double( coordinateComponents(2))
	else
		altOut = 0
	end if

	parseCoordinate = true
	
end function

'----------------------------------------------------------------------------------
' Converts numeric string with "." as decimal separator to double

Dim g_decimalSeparator	' decimal separator in current locale

Function str2double( inStr) ' as double

	if true = IsEmpty(g_decimalSeparator) then
		g_decimalSeparator = Mid( CStr( 1.1), 2, 1)
	end if

	str2double = CDbl( Replace ( inStr, ".", g_decimalSeparator))	

End Function

'----------------------------------------------------------------------------------
' Ruturns file name from URL without extension

Function getFileNameFromUrl( url)

	Dim pos
	Dim fileName

	' remove path from the URL
	pos = InStrRev( url, "/")
	if pos > 0 then
		fileName = Mid( url, pos + 1)
	else
		fileName = url
	end if

	' remove extension from the url
	pos = InStr( fileName, ".")
	if pos > 1 then
		fileName = Mid( fileName, 1, pos - 1)
	else
		fileName = fileName
	end if

	getFileNameFromUrl = fileName

End Function
