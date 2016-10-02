/*----------------------------------------------------------------------------------

(c) SPH Engineering, 2016

CSV file to UgCS route converter

This converter can generate UgCS route from CSV file.

Arguments:
	- CSV full file name
	- UgCS route template full file name

! This converter will work only on Windows.

----------------------------------------------------------------------------------*/

if( 2 != WScript.Arguments.Count()) {
	Log( "Please specify CSV file name and route template in command line.");
	WScript.Quit();
}

csv2route( WScript.Arguments.Item(0), WScript.Arguments.Item(1), 50);

WScript.Quit();

//----------------------------------------------------------------------------------
// Prints string

function Log( msg) {
	WScript.Echo( msg);
}

//----------------------------------------------------------------------------------
// Converts CSV file to UgCS route file

function csv2route( csvFileName, routeTemplateFileName, defaultAltitude)
{
	Log( "CSV file name: " + csvFileName);
	Log( "Template file name: " + routeTemplateFileName);


	// load route template
	var xmlDoc = loadXmlFile( routeTemplateFileName);
	if( null == xmlDoc)
		return;

	// find first waypoint in template 
	var waypointTemplate = xmlDoc.selectSingleNode("//segments");
	if( null == waypointTemplate) {
		Log( "Can't find waypoint in route template!");
		return;
	}

	// load CSV data
	var arrCsvData = loadCsvFile( csvFileName);
	if( 0 == arrCsvData.length) {
		Log( "CSV file doesn't contain any data");
		return;
	}
	else {
		Log( "CSV file contains " + arrCsvData.length + " data lines");
	}

	for( var i = 0; i < arrCsvData.length; i++) {

		// Each data line should contain at leas 2 elements (lattitude and longtitude).
		// Third parameter (lattitude) is optional 

		if( arrCsvData[i].length < 2) {
			Log( "  Line " + i + " contains only " + arrCsvData[i].length + " element(s), skipping");
			continue;
		}
		
		var altitude = defaultAltitude;
		if( 3 <= arrCsvData[i].length) altitude = arrCsvData[i][2];

		addWaypoint( xmlDoc, 
			numberFromCoordinateString( arrCsvData[i][0]), 
			numberFromCoordinateString( arrCsvData[i][1]), 
			altitude, i, waypointTemplate)
		}

	// Remove template waypoint
	waypointTemplate.parentNode.removeChild( waypointTemplate);

	// Save modified UgCS route template as new route 
	var ugcsRouteFileName = getFileName( csvFileName) + ".xml";
	xmlDoc.save( ugcsRouteFileName);
	Log( "UgCS route saved in file '" + ugcsRouteFileName + "'");
}

//----------------------------------------------------------------------------------
// Converts coordinate component string to number, i.e.
//	51.788589N -> 51.788589
//	51.788589S -> -51.788589
//	0.706142E -> 0.706142
//	0.706142w -> -0.706142

function numberFromCoordinateString( coordinateString)
{
	var lastChar = coordinateString.slice( -1);

	var out = parseFloat( coordinateString);

	if( 'N' == lastChar || 'n' == lastChar || 'E' == lastChar || 'e' == lastChar)
		return out;

	return -1 * out;
}

//----------------------------------------------------------------------------------
// Add waypoint to the route
// Return:
//	success: new waypoint (XMLDOMNode)
//	error: null

function addWaypoint( xmlRoute, lattitude, longtitude, altitude, index, prevWaypoint)
{
	var PI = 3.14159265358979;

	// create new waypoint xml structure using previous waypoint as template
	var newWaypoint = prevWaypoint.cloneNode( true);

	// get node with coordinates
	var coordsNode = newWaypoint.selectSingleNode( "figure/ugcs-List/o");
	if( null == coordsNode) {
		Log( "Can't find figure/ugcs-list/o node in waypoint XML");
		return null;
	}

	// Update coordinates

	// Convert longtitude and lattitude to radians
	longtitude = longtitude * PI / 180;
	lattitude = lattitude * PI / 180;

	// Node with waypoint coordinates has following format 
	// <o v7="AGL" v6="0.0" v4="longtitude" v3="lattitude" v2="0"/>

	// We used previous point as template we need to update coordinates and altitude only 
	coordsNode.setAttribute( "v6", altitude);
	coordsNode.setAttribute( "v4", longtitude);
	coordsNode.setAttribute( "v3", lattitude);	
		
	// get node with waypoint order 
	var orderNode = newWaypoint.selectSingleNode( "order");
	if( null == orderNode) {
		Log( "Can't find order node in waypoint XML");
		return null;
	}

	// update waypoint index
	orderNode.setAttribute( "v", index);

	// append new waypoint after previous one
	prevWaypoint.parentNode.appendChild( newWaypoint);
	
		
	// trick - UgCS XML parser can't load XML with more then one element per line
	// so will add text node with new line
	prevWaypoint.parentNode.appendChild( xmlRoute.createTextNode( "\r\n"));
	
}

//----------------------------------------------------------------------------------
// Load XML file and set selection language
// Return:
//	success: XML.DomDocument
//	error: null and displays error message in console


function loadXmlFile( fileName)
{
	var xmlDoc = new ActiveXObject( "Msxml2.DOMDocument");
	xmlDoc.async = false;

	if( false == xmlDoc.load( fileName)) {
		Log( "Can't load XML file " + fileName + ":" + xmlDoc.parseError.reason);
		return null;
	}

	xmlDoc.setProperty( "SelectionLanguage", "XPath");
	
	return xmlDoc;
}

//----------------------------------------------------------------------------------
// Load CSV file
// Return file content in form of array of arrays

function loadCsvFile( fileName)
{
	var FSO = new ActiveXObject("Scripting.FileSystemObject");
	var file = FSO.OpenTextFile( fileName, 1);	// open file for reading
	var content = file.ReadAll();
	file.close();

	return CSVToArray( content);
}

//----------------------------------------------------------------------------------
// ref: http://stackoverflow.com/a/1293163/2343
// This will parse a delimited string into an array of
// arrays. The default delimiter is the comma, but this
// can be overriden in the second argument.

function CSVToArray( strData, strDelimiter )
{

        // Check to see if the delimiter is defined. If not,
        // then default to comma.
        strDelimiter = (strDelimiter || ",");

        // Create a regular expression to parse the CSV values.
        var objPattern = new RegExp(
            (
                // Delimiters.
                "(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +

                // Quoted fields.
                "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +

                // Standard fields.
                "([^\"\\" + strDelimiter + "\\r\\n]*))"
            ),
            "gi"
            );


        // Create an array to hold our data. Give the array
        // a default empty first row.
        var arrData = [[]];

        // Create an array to hold our individual pattern
        // matching groups.
        var arrMatches = null;


        // Keep looping over the regular expression matches
        // until we can no longer find a match.
        while (arrMatches = objPattern.exec( strData )){

            // Get the delimiter that was found.
            var strMatchedDelimiter = arrMatches[ 1 ];

            // Check to see if the given delimiter has a length
            // (is not the start of string) and if it matches
            // field delimiter. If id does not, then we know
            // that this delimiter is a row delimiter.
            if (
                strMatchedDelimiter.length &&
                strMatchedDelimiter !== strDelimiter
                ){

                // Since we have reached a new row of data,
                // add an empty row to our data array.
                arrData.push( [] );

            }

            var strMatchedValue;

            // Now that we have our delimiter out of the way,
            // let's check to see which kind of value we
            // captured (quoted or unquoted).
            if (arrMatches[ 2 ]){

                // We found a quoted value. When we capture
                // this value, unescape any double quotes.
                strMatchedValue = arrMatches[ 2 ].replace(
                    new RegExp( "\"\"", "g" ),
                    "\""
                    );

            } else {

                // We found a non-quoted value.
                strMatchedValue = arrMatches[ 3 ];

            }

            // Now that we have our value string, let's add
            // it to the data array.
            arrData[ arrData.length - 1 ].push( strMatchedValue );
        }

        // Return the parsed data.
        return( arrData );
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
