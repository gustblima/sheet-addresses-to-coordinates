# coding: latin-1
import os
import collections
import configparser
import json
import requests
import sys
import argparse
from xlutils.copy import copy
from xlrd import open_workbook


class SheetsToCoordinates(object):

    def __init__(self, filename, sheet_index):
        self.filename = filename
        self.read_book = open_workbook(filename)
        self.read_sheet = self.read_book.sheet_by_index(sheet_index) 
        self.work_book = copy(self.read_book) 
        self.write_sheet = self.work_book.get_sheet(sheet_index)
        settings = configparser.ConfigParser()
        settings.read("config.ini")
        self.key = settings.get('GoogleApi', 'key')
        self.base_url = "https://maps.googleapis.com/maps/api/geocode/json"
        
    def get_coordinates(self, local):
        """Get latitude and longitude from Google Maps API"""
        params = dict(
            address = local,
            key = self.key
        )
        resp = requests.get(url=self.base_url, params= params)
        data = json.loads(resp.text)
        if data['results'] is not None:
            location = data['results'][0]['geometry']['location']
            return collections.namedtuple('Point', 'lat, lng')(location['lat'],location['lng'])
        return collections.namedtuple('Point', 'lat ,lng')(0,0)

    def convert_rows(self, row_start, col_in, col_out_lat, col_out_lng):
        """Read address from a starting row and column and translate to geocode in another two columns"""
        for row in range(row_start, self.read_sheet.nrows):
            local = self.read_sheet.cell(row, col_in).value
            geo = self.get_coordinates(local)
            print local, geo.lat, geo.lng
            self.write_sheet.write(row, col_out_lat, geo.lat)
            self.write_sheet.write(row, col_out_lng, geo.lng)
        self.write_file()

    def write_file(self):
        """Write output to another file"""
        filename, ext = os.path.splitext(self.filename)
        self.work_book.save(filename + "_output" + ext)
     

def main():
    """Main Method"""
    parser =  argparse.ArgumentParser(description='Transform address to latitude and longitude.')

    parser.add_argument("-f", "--file", help="Input File")
    parser.add_argument("-s", "--sheet",type=int, help="Sheet index")
    parser.add_argument("-r", "--row", type=int, help="Start Row ")
    parser.add_argument("-a", "--address", type=int, help="Column Address")
    parser.add_argument("-lat", "--latitude", type=int, help="Column Latitude")
    parser.add_argument("-lng", "--longitude", type=int, help="Column Longitude")
    args = parser.parse_args()
    if len(sys.argv) < 6:
        print "You need to pass all parameters"
        parser.print_help()
        sys.exit(1)
    a = SheetsToCoordinates(args.file, args.sheet)
    a.convert_rows(args.row, args.address, args.latitude, args.longitude)    

if __name__ == "__main__":
    main()

