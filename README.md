# Address to Coordinates
Reads Address in a sheet, get Latitude and Longitude and write back in another two columns.

# Usage
First you'll need a Google Maps Api Key:
https://developers.google.com/maps/documentation/geocoding/start
After that, put your generated key in the config.ini file
Now you're read to use
# Paramaters
-f *Filename* 

-s *Sheet index*

-r *Start row index*

-a *Address column*

-lat *Latitude Output column*

-lng *Longitude Output Column*

# Example:
>python sheets_to_coordinates.py -f myAddresses.xlsx -s 0 -r 0 -a 1 -lat 2 -lng 3

output will be: myAddresses_output.xlsx
