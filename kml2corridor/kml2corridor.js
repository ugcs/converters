/*----------------------------------------------------------------------------------

(c) SPH Engineering, 2018

KML to UgCS route converter

This converter can generate corridor mapping routes from KML file.
KML files should contain Placemark element with coordinates tag.

Arguments:
	- KML full file name
	- UgCS route template full file name

	- Offset of flight path from centerline (optional, default = 0)

! This converter will work only on Windows.

----------------------------------------------------------------------------------*/

var EATHRADIUS = 6378137;

if (WScript.Arguments.Count() < 2) {
	log("Please specify KML file name and route template in command line.");
	WScript.Quit();
}

var offset = 0;

if (WScript.Arguments.Count() === 3 && WScript.Arguments.Item(2) > 0)
	offset = WScript.Arguments.Item(2);

kmlFile2Routes(WScript.Arguments.Item(0), WScript.Arguments.Item(1), offset);

WScript.Quit();

//----------------------------------------------------------------------------------
// Returns bearing of two points

function getBearing(latitude1, longitude1, latitude2, longitude2) {
	var dLon = this.toRadian(longitude2 - longitude1);
	var y = Math.sin(dLon) * Math.cos(this.toRadian(latitude2));
	var x = Math.cos(this.toRadian(latitude1)) * Math.sin(this.toRadian(latitude2))
		- Math.sin(this.toRadian(latitude1)) * Math.cos(this.toRadian(latitude2)) * Math.cos(dLon);
	var brng = this.toDegrees(Math.atan2(y, x));
	return ((brng + 360) % 360);
}

//----------------------------------------------------------------------------------
// Returns radian value from degrees

function toRadian(value) {
	return value * Math.PI / 180;
}

//----------------------------------------------------------------------------------
// Returns degrees value from radian

function toDegrees(value) {
	return value * 180 / Math.PI;
}

//----------------------------------------------------------------------------------
// Prints string

function log(msg) {
	WScript.Echo(msg);
}

//----------------------------------------------------------------------------------
// Converts KML file to UgCS route files (each Placemark to separate UgCS route file)

function kmlFile2Routes(kmlFileName, routeTemplateFileName, offset)
{
    log("KML file name: " + kmlFileName);
    log("Template file name: " + routeTemplateFileName);

	var xmlDoc = new ActiveXObject("Msxml2.DOMDocument");
	xmlDoc.async = false;

	if(!xmlDoc.load(kmlFileName)) {
        log("Can't load KML file: " + xmlDoc.parseError.reason);
		return;
	}

	xmlDoc.setProperty("SelectionLanguage", "XPath");
	xmlDoc.setProperty("SelectionNamespaces", "xmlns:kml='http://www.opengis.net/kml/2.2'");

	var placemarks = xmlDoc.selectNodes("//kml:Placemark");
	if(placemarks.length === 0) {
        log( "Can't find any Placemark in KML file!");
		return;
	}

	for(var i = 0; i < placemarks.length; ++i) {
		var direction = getDirectionFromStylePlacemark(xmlDoc, placemarks(i));
		placemark2route(placemarks(i), i, kmlFileName, routeTemplateFileName, direction, offset);
	}
}

//----------------------------------------------------------------------------------
// Returns direction of placemark by style color

function getDirectionFromStylePlacemark(xmlDoc, placemark) {
	var styleUrl = placemark.selectSingleNode("kml:styleUrl").text.replace("#", "");
	var styleMaps = xmlDoc.selectNodes("//kml:StyleMap");

	for (var i = 0; i < styleMaps.length; ++i) {
		var styleMapIdAttribute = styleMaps(i).getAttribute("id");

		if (styleMapIdAttribute === styleUrl) {
			var styleMapStyleUrl = styleMaps(i).selectSingleNode("./kml:Pair").selectSingleNode("./kml:styleUrl").text.replace("#", "");
			var styles = xmlDoc.selectNodes("//kml:Style");

			for (var j = 0; j < styles.length; ++j) {
				var styleIdAttribute = styles(j).getAttribute("id");

				if (styleIdAttribute === styleMapStyleUrl) {
                    var color = styles(j).selectSingleNode("./kml:LineStyle").selectSingleNode("./kml:color").text.replace("#", "");
                    return color === "ffff0000" ? 1 : -1; // left = 1, right = -1: see to south or east
				}
			}
            log("Can`t find direction color for placemark.");
            WScript.Quit();
		}
	}
    log("Can`t find style id binding for placemark.");
    WScript.Quit();
}

//----------------------------------------------------------------------------------
// Returns offset longitude and latitude

function calculateOffsetPoint(currentPoint, lastPoint, offset, direction) {
    var currentLongitude = Number(currentPoint[0]);
    var currentLatitude = Number(currentPoint[1]);

    var lastLongitude = Number(lastPoint[0]);
    var lastLatitude = Number(lastPoint[1]);

    var vectorBearing = getBearing(currentLatitude, currentLongitude, lastLatitude, lastLongitude);
    var azimuthPart = getVectorAzimuthPart(vectorBearing);

    var currentLongitudeRadian = toRadian(currentLongitude);
    var currentLatitudeRadian = toRadian(currentLatitude);

    var offsetLatitudeRadian = offset / EATHRADIUS * direction;
    var offsetLongitudeRadian = offset / (EATHRADIUS * Math.cos(currentLatitudeRadian)) * direction;

    // calculating new offset coordinates of point
    if (azimuthPart === 1) {
        currentLongitudeRadian -= offsetLongitudeRadian;
        currentLatitudeRadian += offsetLatitudeRadian;
    } else if (azimuthPart === 2) {
        currentLatitudeRadian += offsetLatitudeRadian;
    } else if (azimuthPart === 3) {
        currentLongitudeRadian += offsetLongitudeRadian;
        currentLatitudeRadian += offsetLatitudeRadian;
    } else if (azimuthPart === 4) {
        currentLongitudeRadian += offsetLongitudeRadian;
        currentLatitudeRadian += offsetLatitudeRadian;
    } else if (azimuthPart === 5) {
        currentLongitudeRadian += offsetLongitudeRadian;
        currentLatitudeRadian -= offsetLatitudeRadian;
    } else if (azimuthPart === 6) {
        currentLatitudeRadian -= offsetLatitudeRadian;
    } else if (azimuthPart === 7) {
        currentLongitudeRadian -= offsetLongitudeRadian;
        currentLatitudeRadian -= offsetLatitudeRadian;
    } else if (azimuthPart === 8) {
        currentLongitudeRadian -= offsetLongitudeRadian;
    }

    var currentLongitudeDegrees = toDegrees(currentLongitudeRadian);
    var currentLatitudeDegrees = toDegrees(currentLatitudeRadian);

    return [currentLongitudeDegrees, currentLatitudeDegrees];
}

//----------------------------------------------------------------------------------
// Converts one KML Placemark to UgCS route

function placemark2route(kmlPlacemark, placemarkIndex, kmlFileName, routeTemplateFileName, direction, offset) {

	var routeName;

	// Try to get KML placemark name - it will be used as UgCS route and file name
	var kmlPolygonName = kmlPlacemark.selectSingleNode("kml:name");

	if(kmlPolygonName === null) {
		routeName = getFileName(kmlFileName) + "." + placemarkIndex;
        log("Found unnamed placemark");
	} else {
		routeName = kmlPolygonName.text;
        log("Found placemark '" + routeName + "'");
	}

	var ugcsRouteTemplate = getUgcsRouteTemplate(routeTemplateFileName);
	if(ugcsRouteTemplate === null)
		return;

	// Update route name in route template
	ugcsRouteTemplate.selectSingleNode("//Route/name").setAttribute("v", routeName);

	// Try to find coordinates node inside of placemark's polygon
	var kmlCoordinates = kmlPlacemark.selectSingleNode(".//kml:coordinates");
	if(kmlCoordinates === null) {
        log("This placemark does not contain any coordinates, exiting...");
		return;
	}

	var cordsArray = kmlCoordinates.text.split(" ");
    if(cordsArray.length < 1) {
        log("Can't found coordinates in placemark '" + routeName + "'!");
        return;
    }

	// fill the straight way path
	for (var i = 0; i < cordsArray.length; ++i) {
    	var iterateDirection = direction;
        var currentPoint = cordsArray[i].split(",");
        var lastPoint;
        if (i === 0) {
        	lastPoint = cordsArray[i + 1].split(",");
        	iterateDirection = -iterateDirection;
        } else {
        	lastPoint = cordsArray[i - 1].split(",");
		}

        // check if have at least longitude and latitude here
        if(currentPoint.length < 2 || lastPoint.length < 2) {
            log("Can't parse coordinates string in placemark '" + routeName + "'!");
            continue;
        }

        var currentOffsetPoint = calculateOffsetPoint(currentPoint, lastPoint, offset, iterateDirection);
		var cordsIndex = direction === 1 ? i : cordsArray.length - i - 1;
        addUgcsFigurePoint(ugcsRouteTemplate, currentOffsetPoint[0], currentOffsetPoint[1], cordsIndex);
	}

	// fill back way path
    for (var j = cordsArray.length - 1; j >= 0; --j) {
    	var iterateRevertDirection = direction;
        var currentRevertPoint = cordsArray[j].split(",");
        var lastRevertPoint;
        if (j === cordsArray.length - 1) {
        	lastRevertPoint = cordsArray[j - 1].split(",");
        	iterateRevertDirection = - iterateRevertDirection;
		} else {
        	lastRevertPoint = cordsArray[j + 1].split(",");
		}

        // check if have at least longitude and latitude here
        if(currentRevertPoint.length < 2 || lastRevertPoint.length < 2) {
            log("Can't parse coordinates string in placemark '" + routeName + "'!");
            continue;
        }

        var currentRevertOffsetPoint = calculateOffsetPoint(currentRevertPoint, lastRevertPoint, offset, iterateRevertDirection);
        var cordsRevertIndex = direction === 1 ? cordsArray.length * 2 - 1 - j: cordsArray.length + j;
        addUgcsFigurePoint(ugcsRouteTemplate, currentRevertOffsetPoint[0], currentRevertOffsetPoint[1], cordsRevertIndex);
    }

	// Save modified UgCS route template as new route 
	var ugcsRouteFileName = routeName + ".xml";

	if(isFileExists(ugcsRouteFileName))	// add KML placemark index to the file name if it's not unique
		ugcsRouteFileName = routeName + " " + placemarkIndex + ".xml";

	var mainSegment = getUgcsSegmentsElement(ugcsRouteTemplate);
	var route = getUgcsRouteElement(ugcsRouteTemplate);
	route.removeChild(mainSegment);
	ugcsRouteTemplate.save(ugcsRouteFileName);
    log("UgCS route saved in file '" + ugcsRouteFileName + "'");
}

// clockwise bearing between points direction part
function getVectorAzimuthPart(bearing) {
	if (bearing >= 0 && bearing < 45)
		return 1;
	if (bearing >= 45 && bearing < 90)
		return 2;
	if (bearing >= 90 && bearing < 135)
		return 3;
    if (bearing >= 135 && bearing < 180)
        return 4;
    if (bearing >= 180 && bearing < 225)
        return 5;
    if (bearing >= 225 && bearing < 270)
        return 6;
    if (bearing >= 270 && bearing < 315)
        return 7;
	return 8;
}

//----------------------------------------------------------------------------------
// Loads UgCS route template from file and clears collection of figure 
// (AreaScan/Perimeter/Photogrammetry tool) points
// Returns:
//	success: route template (Msxml2.DOMDocument)
//	error: null

function getUgcsRouteTemplate(routeTemplateFileName) {

	var xmlDoc = new ActiveXObject("Msxml2.DOMDocument");
	xmlDoc.async = false;
	xmlDoc.preserveWhiteSpace = true;

	if(!xmlDoc.load( routeTemplateFileName)) {
        log("Can't load UgCS route template file '" + routeTemplateFileName + "': " + xmlDoc.parseError.reason);
		return null;
	}

	xmlDoc.setProperty("SelectionLanguage", "XPath");

	// find first figure definition 
	var points = getUgcsFigurePoints(xmlDoc);
	if(points === null) {
        log("Can't find figure definition in template file '" + routeTemplateFileName + "'!");
		return null;
	}
	
	// clear all points from polygon definition
	points.selectNodes("o").removeAll();

	return xmlDoc;
}

//----------------------------------------------------------------------------------
// Returns figure points list node from UgCS route

function getUgcsFigurePoints(ugcsRouteDomDocument) {
	// find first figure definition
	return ugcsRouteDomDocument.selectSingleNode("//Route/segments/figure/ugcs-List");
}

//----------------------------------------------------------------------------------
// Returns segment element from UgCS route

function getUgcsSegmentsElement(ugcsRouteDomDocument) {
	return ugcsRouteDomDocument.selectSingleNode("//Route/segments");
}

//----------------------------------------------------------------------------------
// Returns UgCS route element

function getUgcsRouteElement(ugcsRouteDomDocument) {
	return ugcsRouteDomDocument.selectSingleNode("//Route");
}

//----------------------------------------------------------------------------------
// Adds new point to UgCS route figure definition

function addUgcsFigurePoint(ugcsRouteDomDocument, longtitude, lattitude, coordsIndex) {
	// Convert longitude and latitude to radians
	longtitude = longtitude * Math.PI / 180;
	lattitude = lattitude * Math.PI / 180;

	var route = getUgcsRouteElement(ugcsRouteDomDocument);
	var originalSegment = getUgcsSegmentsElement(ugcsRouteDomDocument);
	var newSegment = originalSegment.cloneNode(true); // deep copy

	newSegment.selectSingleNode("./order").setAttribute("v", coordsIndex);
	newSegment.selectSingleNode("./actionDefinitions").selectSingleNode("./order").setAttribute("v", coordsIndex);

	var ugcsList = newSegment.selectSingleNode("//figure").selectSingleNode("//ugcs-List");

    // create and fill new point
    // it should have format
    // <o v7="AGL" v6="0.0" v4="longitude" v3="latitude" v2="order"/>

    var point = ugcsRouteDomDocument.createElement("o");
    point.setAttribute("v7", "AGL");
    point.setAttribute("v6", "0.0");
    point.setAttribute("v4", longtitude);
    point.setAttribute("v3", lattitude);
    point.setAttribute("v2", "0");

    // trick - UgCS XML parser can't load XML with more then one element per line
    // so will add text node with new line
    ugcsList.appendChild(point);
    ugcsList.appendChild(ugcsRouteDomDocument.createTextNode("\t\n"));

	route.appendChild(newSegment);
	route.appendChild(ugcsRouteDomDocument.createTextNode("\n\t\t"));
}

//----------------------------------------------------------------------------------
// Ruturns file name from full path without extension

function getFileName(path) {

	var fileName;

	// remove path from the URL
	var pos = path.lastIndexOf(path.charAt(path.indexOf(":") + 1));
	if(pos > 0)
		fileName = path.substring(pos + 1);
	else
		fileName = path;

	// remove extension from the URL
	pos = fileName.indexOf(".");
	if(pos > 1)
		fileName = fileName.substring(0, pos);

	return fileName;
}

//----------------------------------------------------------------------------------
// Checks file existence

function isFileExists(fileName) {
	var FSO = new ActiveXObject("Scripting.FileSystemObject");
	return FSO.fileExists(fileName);
}