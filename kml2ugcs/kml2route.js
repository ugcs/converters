/*----------------------------------------------------------------------------------

(c) SPH Engineering, 2017

KML to UgCS route converter

This converter can generate UgCS routes of 3 types (AreaScan, Photogrammetry tool, Perimeter) from KML files.
KML files should contain Placemark element with coordinates tag.

Arguments:
	- KML full file name
	- UgCS route template full file name

! This converter will work only on Windows.

----------------------------------------------------------------------------------*/

if( 2 != WScript.Arguments.Count()) {
	Log( "Please specify KML file name and route template in command line.");
	WScript.Quit();
}

kmlFile2Routes( WScript.Arguments.Item(0), WScript.Arguments.Item(1));

WScript.Quit();

//----------------------------------------------------------------------------------
// Prints string

function Log( msg) {
	WScript.Echo( msg);
}

//----------------------------------------------------------------------------------
// Converts KML file to UgCS route files (each Placemark to separate UgCS route file)

function kmlFile2Routes( kmlFileName, routeTemplateFileName)
{
	Log( "KML file name: " + kmlFileName);
	Log( "Template file name: " + routeTemplateFileName);

	var xmlDoc = new ActiveXObject( "Msxml2.DOMDocument");
	xmlDoc.async = false;

	if( false == xmlDoc.load(kmlFileName)) {
		Log( "Can't load KML file: " + xmlDoc.parseError.reason);
		return;
	}

	xmlDoc.setProperty( "SelectionLanguage", "XPath");
	xmlDoc.setProperty( "SelectionNamespaces", "xmlns:kml='http://www.opengis.net/kml/2.2'");

	var placemarks = xmlDoc.selectNodes("//kml:Placemark");
	if( 0 == placemarks.length) {
		Log( "Can't find any Placemark in KML file!");
		return;
	}

	for( var i = 0; i < placemarks.length; i++)
		placemark2route( placemarks(i), i, kmlFileName, routeTemplateFileName);
}

//----------------------------------------------------------------------------------
// Converts one KML Placemark to UgCS route

function placemark2route( kmlPlacemark, placemarkIndex, kmlFileName, routeTemplateFileName) {

	var routeName;

	// Try to get KML placemark name - it will be used as UgCS route and file name
	var kmlPolygonName = kmlPlacemark.selectSingleNode("kml:name");
	if( null == kmlPolygonName) {
		routeName = getFileName( kmlFileName) + "." + placemarkIndex;
		Log( "Found unnamed placemark");
	}
	else {
		routeName = kmlPolygonName.text;
		Log( "Found placemark '" + routeName + "'");
	}

	var ugcsRouteTemplate = getUgcsRouteTemplate( routeTemplateFileName);
	if( null == ugcsRouteTemplate)
		return;

	// Update route name in route template
	ugcsRouteTemplate.selectSingleNode("//Route/name").setAttribute( "v", routeName);

	// Try to find coordinates node inside of placemark's polygon
	var kmlCoordinates = kmlPlacemark.selectSingleNode(".//kml:coordinates");
	if( null == kmlCoordinates) {
		Log( " this placemark does not contain any coordinates, exiting...");
		return;
	}

	var coordsArray = kmlCoordinates.text.split(" ");

	for( var i = 0; i < coordsArray.length; i++) {

		var coordinateComponents = coordsArray[i].split( ",");

		// check if have at least longtitude and lattitude here
		if( coordinateComponents.length < 2) {
			Log( " Can't parse coordinates string '" + coordsArray[i] + "' in placemark '" + routeName + "'!");			
			continue;
		}

		var longtitude = Number( coordinateComponents[0]);
		var lattitude = Number( coordinateComponents[1]);

		addUgcsFigurePoint( ugcsRouteTemplate, longtitude, lattitude);
	}

	// Save modified UgCS route template as new route 
	var ugcsRouteFileName = routeName + ".xml";

	if( isFileExists(ugcsRouteFileName))	// add KML placemark index to the file name if it's not unique
		ugcsRouteFileName = ugcsRouteFileName + " " + placemarkIndex.toString();

	ugcsRouteTemplate.save( ugcsRouteFileName);
	Log( " UgCS route saved in file '" + ugcsRouteFileName + "'");
}

//----------------------------------------------------------------------------------
// Loads UgCS route template from file and clears collection of figure 
// (AreaScan/Perimeter/Photogrammetry tool) points
// Returns:
//	success: route template (Msxml2.DOMDocument)
//	error: null

function getUgcsRouteTemplate( routeTemplateFileName) {

	var xmlDoc = new ActiveXObject( "Msxml2.DOMDocument");
	xmlDoc.async = false;
	xmlDoc.preserveWhiteSpace = true;

	if( false == xmlDoc.load( routeTemplateFileName)) {
		Log( "Can't load UgCS route template file '" + routeTemplateFileName + "': " + xmlDoc.parseError.reason);
		return null;
	}

	xmlDoc.setProperty( "SelectionLanguage", "XPath");

	// find first figure definition 
	var points = getUgcsFigurePoints( xmlDoc);
	if( null == points) {
		Log( "Can't find figure definition in template file '" + routeTemplateFileName + "'!");
		return null;
	}
	
	// clear all points from polygon definition
	points.selectNodes("o").removeAll();

	return xmlDoc;
}

//----------------------------------------------------------------------------------
// Returns figure points list node from UgCS route

function getUgcsFigurePoints( ugcsRouteDomDocument) {

	// find first figure definition 
	return ugcsRouteDomDocument.selectSingleNode("//Route/segments/figure/ugcs-List");
}

//----------------------------------------------------------------------------------
// Adds new point to UgCS route figure definition

function addUgcsFigurePoint( ugcsRouteDomDocument, longtitude, lattitude) {

	var PI = 3.14159265358979;

	// Convert longtitude and lattitude to radians
	longtitude = longtitude * PI / 180;
	lattitude = lattitude * PI / 180;

	var points = getUgcsFigurePoints( ugcsRouteDomDocument);

	// create and fill new point
	// it should have format 
	// <o v7="AGL" v6="0.0" v4="longtitude" v3="lattitude" v2="order"/>

	var point = ugcsRouteDomDocument.createElement("o");
	point.setAttribute( "v7", "AGL");
	point.setAttribute( "v6", "0.0");
	point.setAttribute( "v4", longtitude);
	point.setAttribute( "v3", lattitude);
	point.setAttribute( "v2", points.selectNodes("o").length);

	points.appendChild( point);

	// trick - UgCS XML parser can't load XML with more then one element per line
	// so will add text node with new line
	points.appendChild( ugcsRouteDomDocument.createTextNode( "\r\n"));
}

//----------------------------------------------------------------------------------
// Ruturns file name from full path without extension

function getFileName( path) {

	var fileName;

	// remove path from the URL
	var pos = path.lastIndexOf( path.charAt( path.indexOf(":") + 1) );
	if( pos > 0)
		fileName = path.substring( pos + 1);
	else
		fileName = path;

	// remove extension from the URL
	pos = fileName.indexOf( ".");
	if( pos > 1)
		fileName = fileName.substring( 0, pos);

	return fileName;
}

//----------------------------------------------------------------------------------
// Checks file existence

function isFileExists( fileName) {

	var FSO = new ActiveXObject("Scripting.FileSystemObject");

	return FSO.fileExists( fileName);
}


