# Country and City Fectcher
A Crawler for Country and City Information using google reverse API and Microsoft Globalization Framework

# Structure:
This API is used to fetch data for a pre-configurated database with info included as:


1- Country Details:


- Country ID (Auto Generated)
- Country Name (String) Note 1
- Country Native Name (String) Note 2
- Country Latin Native Name (String) Note 2


2- City Details:


- City ID (Auto Generated)
- Country Name (String) Note 1
- Country Native Name (String) Note2
- Country Native Latin Name (String) Note 2
- Latitude (decimal) Note 1
- Longitude (decimal) Note 1

# Note 1:
All the fields assosiated with this API are from the following source:

**GeoNames:**
([geonames](https://www.geonames.com) ) 
with the following export:
([geonames./export/dump/cities1500.zip](https://www.geonames.com/export/dump/cities1500.zip) ) 


Convert to excel sheet for easy export using:
Microsoft Excel Library: ([Microsoft.Office.Interop.Excel](https://www.nuget.org/packages/Microsoft.Office.Interop.Excel/) ) 


# Note 2:
All the Data For Cities Included with this API are fetched using:

([Google Maps Reverse API](maps.googleapis.com/maps/api/geocode/xml) ) 
By passing the logitude and the latitiude of the location 
and returining the complete XML output