# -*- coding: utf-8 -*-
"""
Created on Wed Jul 21 14:06:32 2021

@author: CCook
"""
import os
import pandas as pd
import geopandas as gpd
import numpy as np
import xlsxwriter

from copy import deepcopy
from decimal import ROUND_HALF_UP, Decimal, getcontext
from math import *
from pandas import isna
from com.sca.ca.model.ACSDataset import ACSDataset
from com.sca.ca.model.CensusDataset import CensusDataset
from com.sca.ca.model.FacilityList import FacilityList
from com.sca.ca.support.UTM import *

# Describe Demographics Within Range of Specified Facilities
class FacilityProximityAssessment:

    def __init__(self, filename_entry, output_dir, faclist_df, radius, census_df, acs_df):

        # Output path
        self.filename_entry = str(filename_entry)
        self.fullpath = output_dir
        self.faclist_df = faclist_df
        self.censusblks_df = census_df
        self.acs_df = acs_df
        self.formats = None
        self.facility_bin = None
        self.national_bin = None

        # Specify range in km
        self.radius = int(radius)

        # Not GUI-related, but will need to be specified? Uncertain as to the implementation
        # of active_columns?
        self.active_columns = [0, 1, 14, 2, 3, 4, 5, 6, 7, 8, 11, 9, 10, 13]

    def create_formats(self, workbook):
        formats = {}

        formats['top_header'] = workbook.add_format({
            'bold': 1,
            'bottom': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': 1})

        formats['sub_header_1'] = workbook.add_format({
            'bold': 0,
            'bottom': 1,
            'align': 'center',
            'valign': 'bottom',
            'text_wrap': 1})

        formats['sub_header_2'] = workbook.add_format({
            'bold': 0,
            'bottom': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': 1})

        formats['sub_header_3'] = workbook.add_format({
            'bold': 0,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': 1})

        formats['sub_header_4'] = workbook.add_format({
            'bold': 1,
            'align': 'left',
            'valign': 'vcenter'})

        formats['notes'] = workbook.add_format({
            'font_size': 12,
            'bold': 0,
            'align': 'left',
            'valign': 'top',
            'text_wrap': 1})

        formats['number'] = workbook.add_format({
            'num_format': '#,##0'})

        formats['percentage'] = workbook.add_format({
            'num_format': '0.0%'})

        formats['int_percentage'] = workbook.add_format({
            'num_format': '0%'})

        return formats

    def round_to_sigfig(x, sig=1):
        # Convert float to decimal and set rounding definition
        dc = getcontext()
        dc.rounding = ROUND_HALF_UP
        str_x = str(x)
        d = Decimal(str_x)

        if x == 0:
            return 0;

        if isnan(x):
            return float('NaN')

        # Round using decimal definition then switch result back to float
        rounded = round(d, sig-int(floor(log10(abs(x))))-1)
        rounded = float(rounded)
        return rounded

    def append_aggregated_data(self, values, worksheet, formats, startrow):

        data = deepcopy(values)

        # First, select the columns that are relevant
        row_idx = np.array([i for i in range(0, len(data))])
        col_idx = np.array(self.active_columns)
        slice = np.array(data)[row_idx[:, None], col_idx]
        startcol = 1

        numrows = len(slice)
        numcols = len(slice[0])
        for row in range(0, numrows):
            for col in range(0, numcols):

                # total pop kept as raw number, but we're using percentages for the breakdowns...
                value = slice[row][col]
                if value != "":
                    value = float(value)
                    format = formats['percentage'] if value <= 1 else formats['number']
                    worksheet.write_number(startrow+row, startcol+col, value, format)
                else:
                    worksheet.write(startrow+row, startcol+col, value)

        return startrow + numrows

    def tabulate_national_data(self, row):

        population = row['totalpop']
        pct_minority = row['p_minority']
        pct_white = row['pnh_white']
        pct_black = row['pnh_afr_am']
        pct_amerind = row['pnh_am_ind']
        pct_other = row['pnh_othmix']
        pct_hisp = row['pt_hisp']
        pct_age_lt18 = row['p_agelt18']
        pct_age_gt64 = row['p_agegt64']
        edu_universe = row['edu_univ']
        pct_edu_lths = row['p_edulths']
        pov_universe = row['pov_univ']
        pct_lowinc = row['p_2xpov']
        pct_lingiso = row['p_lingiso']
        pct_pov = row['p_pov']

        self.national_bin[0][0] += population
        if not isna(pct_minority):
            self.national_bin[1][1] += pct_white * population
            self.national_bin[0][1] += population
        if not isna(pct_black):
            self.national_bin[1][2] += pct_black * population
            self.national_bin[0][2] += population
        if not isna((pct_amerind)):
            self.national_bin[1][3] += pct_amerind * population
            self.national_bin[0][3] += population
        if not isna(pct_other):
            self.national_bin[1][4] += pct_other * population
            self.national_bin[0][4] += population
        if not isna(pct_hisp):
            self.national_bin[1][5] += pct_hisp * population
            self.national_bin[0][5] += population
        if not isna(pct_age_lt18):
            self.national_bin[1][6] += pct_age_lt18 * population
            self.national_bin[0][6] += population
        if not isna(pct_age_gt64):
            self.national_bin[1][8] += pct_age_gt64 * population
            self.national_bin[0][8] += population
        if not isna(pct_age_lt18) and not isna(pct_age_gt64):
            self.national_bin[1][7] += (100 - pct_age_gt64 - pct_age_lt18) * population
            self.national_bin[0][7] += population
        if not isna(edu_universe):
            self.national_bin[1][9] += edu_universe * 100
            self.national_bin[0][9] += population
        if not isna(pov_universe):
            self.national_bin[1][15] += pov_universe * 100
            self.national_bin[0][15] += population
        if not isna(edu_universe) and not isna(pct_edu_lths):
            self.national_bin[1][10] += pct_edu_lths * edu_universe
            self.national_bin[0][10] += edu_universe
        if not isna(pov_universe):
            self.national_bin[1][11] += pct_pov * pov_universe
            self.national_bin[0][11] += pov_universe
        if not isna(pov_universe) and not isna(pct_lowinc):
            self.national_bin[1][12] += pct_lowinc * pov_universe
            self.national_bin[0][12] += pov_universe
        if not isna(pct_lingiso):
            self.national_bin[1][13] += pct_lingiso * population
            self.national_bin[0][13] += population
        if not isna(pct_minority):
            self.national_bin[1][14] += pct_minority * population
            self.national_bin[0][14] += population

    def tabulate_facility_data(self, row):

        population = row['population']
        pct_minority = row['p_minority']
        pct_white = row['pnh_white']
        pct_black = row['pnh_afr_am']
        pct_amerind = row['pnh_am_ind']
        pct_other = row['pnh_othmix']
        pct_hisp = row['pt_hisp']
        pct_age_lt18 = row['p_agelt18']
        pct_age_gt64 = row['p_agegt64']
        edu_universe = row['edu_univ']
        pct_edu_lths = row['p_edulths']
        pov_universe = row['pov_univ']
        pct_lowinc = row['p_2xpov']
        pct_lingiso = row['p_lingiso']
        pct_pov = row['p_pov']
        total_pop = row['totalpop']

        self.facility_bin[0][0] += population
        if not isna(pct_minority):
            self.facility_bin[1][1] += pct_white * population
            self.facility_bin[0][1] += population
        if not isna(pct_black):
            self.facility_bin[1][2] += pct_black * population
            self.facility_bin[0][2] += population
        if not isna((pct_amerind)):
            self.facility_bin[1][3] += pct_amerind * population
            self.facility_bin[0][3] += population
        if not isna(pct_other):
            self.facility_bin[1][4] += pct_other * population
            self.facility_bin[0][4] += population
        if not isna(pct_hisp):
            self.facility_bin[1][5] += pct_hisp * population
            self.facility_bin[0][5] += population
        if not isna(pct_age_lt18):
            self.facility_bin[1][6] += pct_age_lt18 * population
            self.facility_bin[0][6] += population
        if not isna(pct_age_gt64):
            self.facility_bin[1][8] += pct_age_gt64 * population
            self.facility_bin[0][8] += population
        if not isna(pct_age_lt18) and not isna(pct_age_gt64):
            self.facility_bin[1][7] += (100 - pct_age_gt64 - pct_age_lt18) * population
            self.facility_bin[0][7] += population
        if not isna(edu_universe):
            self.facility_bin[1][9] += (edu_universe/total_pop * population) * 100
            self.facility_bin[0][9] += population
        if not isna(pov_universe):
            self.facility_bin[1][15] += (pov_universe/total_pop * population) * 100
            self.facility_bin[0][15] += population
        if not isna(edu_universe) and not isna(pct_edu_lths):
            self.facility_bin[1][10] += pct_edu_lths * (edu_universe/total_pop * population)
            self.facility_bin[0][10] += edu_universe
        if not isna(pov_universe):
            self.facility_bin[1][11] += pct_pov * (pov_universe/total_pop * population)
            self.facility_bin[0][11] += pov_universe
        if not isna(pov_universe) and not isna(pct_lowinc):
            self.facility_bin[1][12] += pct_lowinc * (pov_universe/total_pop * population)
            self.facility_bin[0][12] += pov_universe
        if not isna(pct_lingiso):
            self.facility_bin[1][13] += pct_lingiso * population
            self.facility_bin[0][13] += population
        if not isna(pct_minority):
            self.facility_bin[1][14] += pct_minority * population
            self.facility_bin[0][14] += population

    def create(self):
        self.create_workbook()
        self.calculate_distances()
        self.close_workbook()

    # Distance calculation
    # This utilizes geopandas rather than the query function used in HEM4
    # As distances will need to be calculated for each facility there are many coordinate pairs,
    # which go far faster in this method (~5 min per facility)
    # than if iterated pairwise using just coordinates (~25 min per facility)
    # Still need to develop a way to keep distances linked to facilities for bin creation and output
    def calculate_distances(self):

        start_row = 3

        # Create national bin and tabulate population weighted demographic stats for each sub group.
        self.national_bin = [[0]*16 for _ in range(2)]
        self.acs_df.apply(lambda row: self.tabulate_national_data(row), axis=1)

        # Calculate averages by dividing population for each sub group
        for index in range(1, 16):
            if index == 11:
                self.national_bin[1][index] = self.national_bin[1][index] / (100 * self.national_bin[0][0])
            else:
                self.national_bin[1][index] = self.national_bin[1][index] / (100 * self.national_bin[0][index])

        self.national_bin[0][15] = self.national_bin[0][0] * self.national_bin[1][15]
        for index in range(1, 15):
            if index == 10:
                self.national_bin[0][index] = self.national_bin[0][9] * self.national_bin[1][index]
            else:
                self.national_bin[0][index] = self.national_bin[0][0] * self.national_bin[1][index]

        self.national_bin[1][0] = ""
        start_row = self.append_aggregated_data(
            self.national_bin, self.worksheet, self.formats, start_row) + 1


        for index, row in self.faclist_df.iterrows():
            
            print('Calculating distances for ' + self.faclist_df['facility_id'][index])
            
            self.facility_bin = [[0]*16 for _ in range(2)]
            
            fac_lat = row['lat']
            fac_lon = row['lon']
            fac_latrad = radians(row['lat'])
            fac_lonrad = radians(row['lon'])

            # Convert this facility's lat/lon to UTM
            fac_utmn, fac_utme, fac_utmz, hemi, epsg = UTM.ll2utm(fac_lat, fac_lon)
            
            # Create geodataframe of this one facility
            latlon = [[fac_lat, fac_lon]]
            fac_df = pd.DataFrame(latlon, columns=['lat', 'lon'])
            fac_gdf = gpd.GeoDataFrame(
                fac_df, geometry=gpd.points_from_xy(
                fac_df.lon, fac_df.lat, crs='epsg:4269'))
            fac_gdf = fac_gdf.to_crs(epsg)
            
           
            # Subset census DF to one latitude above and one below and one longitude
            # west and east of this facility
            census_box = self.censusblks_df[(self.censusblks_df['lat'] >= fac_lat-1)
                                                & (self.censusblks_df['lat'] <= fac_lat+1)
                                                & (self.censusblks_df['lon'] >= fac_lon-1)
                                                & (self.censusblks_df['lon'] <= fac_lon+1)]
            
            # Create geodataframe of census_latband and census_lonband and then convert CRS to UTM of facility
            censusblks_gdf = gpd.GeoDataFrame(
                census_box, geometry=gpd.points_from_xy(
                census_box.lon, census_box.lat, crs='epsg:4269'))
            censusblks_gdf = censusblks_gdf.to_crs(epsg)
            
            censusblks_gdf['utme'] = censusblks_gdf.geometry.x
            censusblks_gdf['utmn'] = censusblks_gdf.geometry.y
            
            # Compute distance between blocks and facility (in meters)
            censusblks_gdf['dist_m'] = censusblks_gdf.apply(lambda row: np.sqrt((fac_utme - row['utme'])**2 +
                                        (fac_utmn - row['utmn'])**2), axis=1)
            
            # Subset to user defined radius
            blksinrange_gdf = censusblks_gdf[censusblks_gdf['dist_m'] <= self.radius*1000]

            # Remove blocks corresponding to schools, monitors, etc.
            blksinrange_gdf = blksinrange_gdf.loc[
                (~blksinrange_gdf['blkid'].str.contains('S')) &
                (~blksinrange_gdf['blkid'].str.contains('M'))]

            blksinrange_gdf['bkgrp'] = blksinrange_gdf['blkid'].astype(str).str[:12]
                        
            # Merge with ACS
            # Note: Current merge only picks up numeric census blk ids, some blkids have other
            # characters not caught in ACS file.
            acsinrange_gdf = blksinrange_gdf.merge(
                self.acs_df.astype({'bkgrp': 'str'}), left_on='bkgrp', right_on='bkgrp')
            
            acs_columns = ['population', 'totalpop', 'p_minority', 'pnh_white', 'pnh_afr_am',
                           'pnh_am_ind', 'pnh_othmix', 'pt_hisp', 'p_agelt18', 'p_agegt64',
                           'p_2xpov', 'p_pov', 'age_25up', 'p_edulths', 'p_lingiso',
                           'age_univ', 'pov_univ', 'edu_univ', 'iso_univ', 'pov_fl', 'iso_fl']
            acsinrange_df = pd.DataFrame(acsinrange_gdf, columns=acs_columns)
        
            # Create facility bin and tabulate population weighted demographic stats for each sub
            # group.
            acsinrange_df.apply(lambda row: self.tabulate_facility_data(row), axis=1)
                    
            # Calculate averages by dividing population for each sub group
            for col_index in range(1, 16):
                if col_index == 11:
                    if (100 * self.facility_bin[0][col_index]) == 0:
                        self.facility_bin[1][col_index] = 0
                    else:
                        self.facility_bin[1][col_index] = self.facility_bin[1][col_index] / (100 * self.facility_bin[0][0])
                else:
                    if (100 * self.facility_bin[0][col_index]) == 0:
                        self.facility_bin[1][col_index] = 0
                    else:
                        self.facility_bin[1][col_index] = self.facility_bin[1][col_index] / (100 * self.facility_bin[0][col_index])
        
            self.facility_bin[0][15] = self.facility_bin[0][0] * self.facility_bin[1][15]
            for col_index in range(1, 15):
                if col_index == 10:
                    self.facility_bin[0][col_index] = self.facility_bin[0][9] * self.facility_bin[1][col_index]
                else:
                    self.facility_bin[0][col_index] = self.facility_bin[0][0] * self.facility_bin[1][col_index]
        
            self.facility_bin[1][0] = ""

            start_row = self.append_aggregated_data(
                self.facility_bin, self.worksheet, self.formats, start_row)

    # Create Workbook
    # Final workbook should have similar formatting as ej tables, with two rows for nationwide
    # demographics (population and percentages) and two rows for each facility provided in the
    # original faclist. Facility names should also be provided in column A, although that has not
    # yet been added.
    def create_workbook(self):
        output_dir = self.fullpath
        if not (os.path.exists(output_dir) or os.path.isdir(output_dir)):
            os.mkdir(output_dir)
        filename = os.path.join(output_dir, self.filename_entry + '.xlsx')
        tablename = 'Population Demographics within ' + str(self.radius) + ' km of Source Facilities'
        self.workbook = xlsxwriter.Workbook(filename)
        self.worksheet = self.workbook.add_worksheet('Facility Demographics')
        self.formats = self.create_formats(self.workbook)

        column_headers = ['Total Population', 'White', 'Minority\u1D9C', 'African American',
                          'Native American', 'Other and Multiracial', 'Hispanic or Latino\u1D48',
                          'Age (Years)\n0-17', 'Age (Years)\n18-64', 'Age (Years)\n>=65',
                          'People Living Below the Poverty Level', 'Total Number >= 25 Years Old',
                          'Number >= 25 Years Old without a High School Diploma',
                          'People Living in Linguistic Isolation']

        firstcol = 'A'
        lastcol = chr(ord(firstcol) + len(column_headers))
        top_header_coords = firstcol+'1:'+lastcol+'1'

        # Increase the cell size of the merged cells to highlight the formatting.
        self.worksheet.set_column(top_header_coords, 12)
        self.worksheet.set_row(0, 30)

        # Create top level header
        self.worksheet.merge_range(top_header_coords, tablename, self.formats['top_header'])

        # Create column headers
        self.worksheet.merge_range("A2:A3", 'Population Basis', self.formats['sub_header_2'])
        self.worksheet.merge_range("A4:A5", 'Nationwide', self.formats['sub_header_2'])
        self.worksheet.merge_range("B2:N2", 'Demographic Group',  self.formats['sub_header_3'])

        self.worksheet.set_row(2, 72, self.formats['sub_header_2'])
        for col_num, data in enumerate(column_headers):
            self.worksheet.write(2, col_num+1, data)

        col = 'B'
        for header in column_headers:
            header_coords = col+'4:'+col+'5'
            self.worksheet.merge_range(header_coords, header, self.formats['sub_header_3'])
            self.worksheet.set_column(top_header_coords, 12)
            col = chr(ord(col) + 1)

        # Add Facility Names
        facname_list = self.faclist_df['facility_id'].tolist()
        row_num = 6
        for index, data in enumerate(facname_list):
            self.worksheet.merge_range(row_num, 0, row_num + 1, 0, data, self.formats['sub_header_3'])
            row_num = row_num + 2

        last_data_row = 2 * len(facname_list) + 10

        # Create notes
        first_notes_row = last_data_row + 1
        last_notes_row = first_notes_row + 4
        firstcol = 'A'
        lastcol = chr(ord(firstcol) + len(column_headers))
        notes_coords = firstcol+str(first_notes_row)+':'+lastcol+str(last_notes_row)
        self.worksheet.merge_range(notes_coords, 'Notes:\n\n' + \
          '\u1D43Total nationwide population includes all 50 states plus Puerto Rico. ' + \
          '\nDistributions by race, ethnicity, age, education, income and linguistic isolation are based on ' + \
          'demographic information at the census block group level.\n' + \
          '\u1D9CThe minority population includes people identifying as African American, Native American, Other ' + \
          'and Multiracial, or Hispanic/Latino. Measures are taken to avoid double counting of people identifying ' + \
          'as both Hispanic/Latino and a racial minority.\n' + \
          '\u1D48In order to avoid double counting, the "Hispanic or Latino" category is treated as a distinct ' + \
          'demographic category for these analyses. A person is identified as one of five racial/ethnic ' + \
          'categories above: White, African American, Native American, Other and Multiracial, or Hispanic/Latino.\n' + \
          '\u1D49The population-weighted average risk takes into account risk levels at all. ',  self.formats['notes'])

    def close_workbook(self):
        self.workbook.close()
