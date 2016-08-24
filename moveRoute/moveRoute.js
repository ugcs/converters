/*----------------------------------------------------------------------------------

(c) SPH Engineering, 2016

This script moves UgCS route or mission

Arguments:
	- UgCS route or mission full file name
	- new lattitude for first route or mission waypoint
	- new longtitude for first route or mission waypoint

Coordinates should be in decimal format

Example:

	cscript moveRoute.js myroute.xml 22.42929880348565 114.1145629706805 

! This script will work only on Windows.

----------------------------------------------------------------------------------*/

if( 3 != WScript.Arguments.Count()) {
	Log( "Please specify UgCS route or mission file name, lattitude, longtitude in command line.");
	WScript.Quit();
}

moveRoute( WScript.Arguments.Item(0), WScript.Arguments.Item(1), WScript.Arguments.Item(2));

WScript.Quit();

//----------------------------------------------------------------------------------
// Prints string

function Log( msg) {
	WScript.Echo( msg);
}

//----------------------------------------------------------------------------------
// Moves UgCS route or mission to specified for first waypoint coordinates

function moveRoute( routeFileName, lattitude, longtitude)
{
	Log( "Input file name: " + routeFileName);
	Log( "Lattitude: " + lattitude);
	Log( "Longtitude: " + longtitude);

	var xmlDoc = new ActiveXObject( "Msxml2.DOMDocument");
	xmlDoc.async = false;

	if( false == xmlDoc.load(routeFileName)) {
		Log( "Can't load file: " + xmlDoc.parseError.reason);
		return;
	}

	xmlDoc.setProperty( "SelectionLanguage", "XPath");

	// Select all waypoints from route 
	var points = xmlDoc.selectNodes("//ugcs-List[@type='FigurePoint']/o");
	if( 0 == points.length) {
		Log( "Can't find any waypoint in file!");
		return;
	}

	// Waypoints format: 
	// <o ... v3="lattitude" v4="longtitude" ... />

	// UgCS uses radians for longtitude and lattitude
	// Convert longtitude and lattitude to radians
	var PI = 3.14159265358979;
	lattitude = Number(lattitude) * PI / 180;
	longtitude = Number(longtitude) * PI / 180;

	// Calculate offsets
	var lattitudeOffset = lattitude - Number(points(0).getAttribute("v3"));
	var longtitudeOffset = longtitude - Number(points(0).getAttribute("v4"));

	Log( "Offsets:")
	Log( " lattitude: " + lattitudeOffset);
	Log( " longtitude: " + longtitudeOffset);

	// Move all waypoints
	for( var i = 0; i < points.length; i++) {
		points(i).setAttribute("v3", lattitudeOffset + Number(points(i).getAttribute("v3")));
		points(i).setAttribute("v4", longtitudeOffset + Number(points(i).getAttribute("v4")));
	}
		
	// Save modified UgCS route/mission file
	routeFileName = routeFileName.replace(".xml", "-moved.xml"); 
	xmlDoc.save( routeFileName);
	Log( " UgCS route saved in file '" + routeFileName + "'");
}
