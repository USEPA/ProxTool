# -*- coding: utf-8 -*-
"""
Created on Wed Jul 21 14:06:32 2021

@author: CCook
"""
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
import os
from pyproj import _datadir, datadir
from fiona import _shim, schema
import geopandas as gpd
import numpy as np
import xlsxwriter
import pandas as pd

from copy import deepcopy
from decimal import ROUND_HALF_UP, Decimal, getcontext
from pandas import isna
# from tkinter import messagebox
# from com.sca.ca.model.ACSDataset import ACSDataset
# from com.sca.ca.model.CensusDataset import CensusDataset
# from com.sca.ca.model.FacilityList import FacilityList
from com.sca.ca.support.UTM import *


# Describe Demographics Within Range of Specified Facilities
class FacilityProximityAssessment:

    def __init__(self, filename_entry, output_dir, faclist_df, radius, census_df, acs_df, 
                 acsCountyTract_df):

        # Output path
        self.filename_entry = str(filename_entry)
        self.fullpath = output_dir
        self.faclist_df = faclist_df
        self.censusblks_df = census_df
        self.acs_df = acs_df
        self.acsCountyTract_df = acsCountyTract_df
        self.formats = None
        self.facility_bin = None
        self.national_bin = None
        self.rungroup_bin = None
        
        # Initialize list of used blocks
        self.used_blocks = []
        
        # Initialize set to hold missing blockgroups
        self.missingbkgrps = set()

        # Specify range in km
        self.radius = int(radius)

        # Identify the relevant column indexes from the national and facility bins
        self.active_columns = [0, 1, 14, 2, 3, 4, 5, 6, 7, 8, 11, 9, 10, 13]
        
        # Needed columns from census block dataframe
        self.neededBlockColumns = ['blkid', 'population', 'lat', 'lon']
        

    def haversineDistance(self, lon1, lat1, lon2, lat2):
        """
        Calculate the great circle distance in kilometers between two points 
        on the earth (specified in decimal degrees)
        """
        # convert decimal degrees to radians 
        lon1, lat1, lon2, lat2 = map(np.deg2rad, [lon1, lat1, lon2, lat2])
        
        # haversine formula 
        dlon = lon2 - lon1 
        dlat = lat2 - lat1 
        a = np.sin(dlat/2)**2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon/2)**2
        c = 2 * np.arcsin(np.sqrt(a)) 
        r = 6371 # Radius of earth in kilometers. Use 3956 for miles. Determines return value units.
        return c * r        
 
       
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
            'font_size': 11,
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
                    format = formats['percentage'] if row == 1 else formats['number']
                    worksheet.write_number(startrow+row, startcol+col, value, format)
                else:
                    worksheet.write(startrow+row, startcol+col, value)

        return startrow + numrows

    
    def tabulate_rungroup_data(self, df):
        
        rungroup_df = df[~df['blkid'].isin(self.used_blocks)]
        
        self.rungroup_bin[0][0] += rungroup_df['population'].sum()
        self.rungroup_bin[0][1] += rungroup_df[rungroup_df['p_minority'].notna()]['population'].sum()
        self.rungroup_bin[0][2] += rungroup_df[rungroup_df['pnh_afr_am'].notna()]['population'].sum()
        self.rungroup_bin[0][3] += rungroup_df[rungroup_df['amerind'].notna()]['population'].sum()
        self.rungroup_bin[0][4] += rungroup_df[rungroup_df['pnh_othmix'].notna()]['population'].sum()
        self.rungroup_bin[0][5] += rungroup_df[rungroup_df['pt_hisp'].notna()]['population'].sum()
        self.rungroup_bin[0][6] += rungroup_df[rungroup_df['p_agelt18'].notna()]['population'].sum()
        self.rungroup_bin[0][7] += rungroup_df[(rungroup_df['p_agelt18'].notna()) &
                                              (rungroup_df['p_agegt64'].notna())]['population'].sum()
        self.rungroup_bin[0][8] += rungroup_df[rungroup_df['p_agegt64'].notna()]['population'].sum()
        self.rungroup_bin[0][9] += rungroup_df[rungroup_df['edu_univ'].notna()]['population'].sum()
        self.rungroup_bin[0][10] += rungroup_df[(rungroup_df['edu_univ'].notna()) &
                                   (rungroup_df['p_edulths'].notna())]['eduuniv'].sum()
        self.rungroup_bin[0][11] += rungroup_df[rungroup_df['pov_univ'].notna()]['population'].sum()
        self.rungroup_bin[0][12] += rungroup_df[(rungroup_df['pov_univ'].notna()) &
                                   (rungroup_df['p_2xpov'].notna())]['pov_univ'].sum()
        self.rungroup_bin[0][13] += rungroup_df[rungroup_df['p_lingiso'].notna()]['population'].sum()
        self.rungroup_bin[0][14] += rungroup_df[rungroup_df['p_minority'].notna()]['population'].sum()
        self.rungroup_bin[0][15] += rungroup_df[rungroup_df['pov_univ'].notna()]['population'].sum()
        
        self.rungroup_bin[1][1] += rungroup_df[rungroup_df['white'].notna()]['white'].sum()
        self.rungroup_bin[1][2] += rungroup_df[rungroup_df['black'].notna()]['black'].sum()
        self.rungroup_bin[1][3] += rungroup_df[rungroup_df['amerind'].notna()]['amerind'].sum()
        self.rungroup_bin[1][4] += rungroup_df[rungroup_df['other'].notna()]['other'].sum()
        self.rungroup_bin[1][5] += rungroup_df[rungroup_df['hisp'].notna()]['hisp'].sum()
        self.rungroup_bin[1][6] += rungroup_df[rungroup_df['agelt18'].notna()]['agelt18'].sum()
        self.rungroup_bin[1][7] += rungroup_df[rungroup_df['age18to64'].notna()]['age18to64'].sum()
        self.rungroup_bin[1][8] += rungroup_df[rungroup_df['agegt64'].notna()]['agegt64'].sum()  
        self.rungroup_bin[1][9] += rungroup_df[rungroup_df['eduuniv100'].notna()]['eduuniv100'].sum()
        self.rungroup_bin[1][10] += rungroup_df[rungroup_df['nohs'].notna()]['nohs'].sum()
        self.rungroup_bin[1][11] += rungroup_df[rungroup_df['pov'].notna()]['pov'].sum()
        self.rungroup_bin[1][12] += rungroup_df[rungroup_df['pov2x'].notna()]['pov2x'].sum()
        self.rungroup_bin[1][13] += rungroup_df[rungroup_df['lingiso'].notna()]['lingiso'].sum()
        self.rungroup_bin[1][14] += rungroup_df[rungroup_df['minority'].notna()]['minority'].sum()
        self.rungroup_bin[1][15] += rungroup_df[rungroup_df['povuniv100'].notna()]['povuniv100'].sum()            
        



    def create(self):
        self.create_workbook()
        self.calculate_distances()
        self.close_workbook()
                
        # Write out any missing blockgroups
        if len(self.missingbkgrps) > 0:
            missfname = self.filename_entry + '_' + 'missing_block_groups' + '_' + str(self.radius) + 'km.txt'
            misspath = os.path.join(self.fullpath, missfname)
            
            with open(misspath, 'w') as f:
                for item in self.missingbkgrps:
                    f.write("%s\n" % item)
        

    def calculate_distances(self):

        # Distance calculation
        # This utilizes geopandas rather than the query function used in HEM4.
        # As distances will need to be calculated for each facility there are many coordinate pairs,
        # which go far faster in this method than if iterated pairwise using just coordinates.

        # Initialize starting data rows for the facility and sortable sheets
        start_row = 3
        sort_row = 1

        #------------------------------------------------------------------------------------------
        # Create national bin and tabulate population weighted demographic stats for each sub group.
        #------------------------------------------------------------------------------------------
        self.national_bin = [[0]*16 for _ in range(2)]

        national_acs = self.acs_df
                
        national_acs['white'] = national_acs['pnh_white'] * national_acs['totalpop']
        national_acs['black'] = national_acs['pnh_afr_am'] * national_acs['totalpop']
        national_acs['amerind'] = national_acs['pnh_am_ind'] * national_acs['totalpop']
        national_acs['other'] = national_acs['pnh_othmix'] * national_acs['totalpop']
        national_acs['hisp'] = national_acs['pt_hisp'] * national_acs['totalpop']
        national_acs['agelt18'] = national_acs['p_agelt18'] * national_acs['totalpop']
        national_acs['agegt64'] = national_acs['p_agegt64'] * national_acs['totalpop']
        national_acs['age18to64'] = (100 - national_acs['p_agelt18'] - national_acs['p_agegt64']) * national_acs['totalpop']
        national_acs['eduuniv100'] = national_acs['edu_univ'] * 100 
        national_acs['povuniv100'] = national_acs['pov_univ'] * 100 
        national_acs['nohs'] = national_acs['p_edulths'] * national_acs['edu_univ']
        national_acs['pov'] = national_acs['p_pov'] * national_acs['pov_univ']
        national_acs['pov2x'] = national_acs['p_2xpov'] * national_acs['pov_univ']
        national_acs['lingiso'] = national_acs['p_lingiso'] * national_acs['totalpop']
        national_acs['minority'] = national_acs['p_minority'] * national_acs['totalpop']
        
        self.national_bin[0][0] = national_acs['totalpop'].sum()
        self.national_bin[0][1] = national_acs[national_acs['p_minority'].notna()]['totalpop'].sum()
        self.national_bin[0][2] = national_acs[national_acs['pnh_afr_am'].notna()]['totalpop'].sum()
        self.national_bin[0][3] = national_acs[national_acs['amerind'].notna()]['totalpop'].sum()
        self.national_bin[0][4] = national_acs[national_acs['pnh_othmix'].notna()]['totalpop'].sum()
        self.national_bin[0][5] = national_acs[national_acs['pt_hisp'].notna()]['totalpop'].sum()
        self.national_bin[0][6] = national_acs[national_acs['p_agelt18'].notna()]['totalpop'].sum()
        self.national_bin[0][7] = national_acs[(national_acs['p_agelt18'].notna()) &
                                              (national_acs['p_agegt64'].notna())]['totalpop'].sum()
        self.national_bin[0][8] = national_acs[national_acs['p_agegt64'].notna()]['totalpop'].sum()
        self.national_bin[0][9] = national_acs[national_acs['edu_univ'].notna()]['totalpop'].sum()
        self.national_bin[0][10] = national_acs[(national_acs['edu_univ'].notna()) &
                                   (national_acs['p_edulths'].notna())]['edu_univ'].sum()
        self.national_bin[0][11] = national_acs[national_acs['pov_univ'].notna()]['totalpop'].sum()
        self.national_bin[0][12] = national_acs[(national_acs['pov_univ'].notna()) &
                                   (national_acs['p_2xpov'].notna())]['pov_univ'].sum()
        self.national_bin[0][13] = national_acs[national_acs['p_lingiso'].notna()]['totalpop'].sum()
        self.national_bin[0][14] = national_acs[national_acs['p_minority'].notna()]['totalpop'].sum()
        self.national_bin[0][15] = national_acs[national_acs['pov_univ'].notna()]['totalpop'].sum()

        self.national_bin[1][1] = national_acs[national_acs['white'].notna()]['white'].sum()
        self.national_bin[1][2] = national_acs[national_acs['black'].notna()]['black'].sum()
        self.national_bin[1][3] = national_acs[national_acs['amerind'].notna()]['amerind'].sum()
        self.national_bin[1][4] = national_acs[national_acs['other'].notna()]['other'].sum()
        self.national_bin[1][5] = national_acs[national_acs['hisp'].notna()]['hisp'].sum()
        self.national_bin[1][6] = national_acs[national_acs['agelt18'].notna()]['agelt18'].sum()
        self.national_bin[1][7] = national_acs[national_acs['age18to64'].notna()]['age18to64'].sum()
        self.national_bin[1][8] = national_acs[national_acs['agegt64'].notna()]['agegt64'].sum()  
        self.national_bin[1][9] = national_acs[national_acs['eduuniv100'].notna()]['eduuniv100'].sum()
        self.national_bin[1][10] = national_acs[national_acs['nohs'].notna()]['nohs'].sum()
        self.national_bin[1][11] = national_acs[national_acs['pov'].notna()]['pov'].sum()
        self.national_bin[1][12] = national_acs[national_acs['pov2x'].notna()]['pov2x'].sum()
        self.national_bin[1][13] = national_acs[national_acs['lingiso'].notna()]['lingiso'].sum()
        self.national_bin[1][14] = national_acs[national_acs['minority'].notna()]['minority'].sum()
        self.national_bin[1][15] = national_acs[national_acs['povuniv100'].notna()]['povuniv100'].sum()
        
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
        
        # Write to facility sheet
        start_row = self.append_aggregated_data(
            self.national_bin, self.worksheet_facility, self.formats, start_row) + 1


        
        # Process each facility
        for index, row in self.faclist_df.iterrows():
            
            print('Calculating proximity for facility: ' + self.faclist_df['facility_id'][index])
                            
            self.facility_bin = [[0]*16 for _ in range(2)]
            
            fac_lat = row['lat']
            fac_lon = row['lon']
            # fac_latrad = radians(row['lat'])
            # fac_lonrad = radians(row['lon'])

            # # Convert this facility's lat/lon to UTM
            # fac_utmn, fac_utme, fac_utmz, hemi, epsg = UTM.ll2utm(fac_lat, fac_lon)
                        
            # # Create geodataframe of this one facility
            # latlon = [[fac_lat, fac_lon]]
            # fac_df = pd.DataFrame(latlon, columns=['lat', 'lon'])
            # fac_gdf = gpd.GeoDataFrame(
            #     fac_df, geometry=gpd.points_from_xy(
            #     fac_df.lon, fac_df.lat, crs='epsg:4269'))
            # fac_gdf = fac_gdf.to_crs(epsg)
            
            # Subset census DF to half latitude above and half below and one longitude
            # west and east of this facility
            census_box = self.censusblks_df[(self.censusblks_df['lat'] >= fac_lat-0.5)
                                                & (self.censusblks_df['lat'] <= fac_lat+0.5)
                                                & (self.censusblks_df['lon'] >= fac_lon-1)
                                                & (self.censusblks_df['lon'] <= fac_lon+1)]
            
            # Reduce the number of columns
            census_box = census_box[self.neededBlockColumns]
            
            # Compute distance in km between each census block and the facility
            census_box['dist_km'] = self.haversineDistance(fac_lon, fac_lat, census_box['lon'], census_box['lat'])
            
            # Subset census blocks to user defined radius
            blksinrange_df = census_box[census_box['dist_km'] <= self.radius]
            
            # # Create geodataframe of census_latband and census_lonband and then convert CRS to UTM of facility
            # censusblks_gdf = gpd.GeoDataFrame(
            #     census_box, geometry=gpd.points_from_xy(
            #     census_box.lon, census_box.lat, crs='epsg:4269'))
            # censusblks_gdf = censusblks_gdf.to_crs(epsg)
            
            # censusblks_gdf['utme'] = censusblks_gdf.geometry.x
            # censusblks_gdf['utmn'] = censusblks_gdf.geometry.y
            
            # # Compute distance between blocks and facility (in meters)
            # censusblks_gdf['dist_m'] = censusblks_gdf.apply(lambda row: np.sqrt((fac_utme - row['utme'])**2 +
            #                             (fac_utmn - row['utmn'])**2), axis=1)
                        
            # # Subset to user defined radius
            # blksinrange_gdf = censusblks_gdf[censusblks_gdf['dist_m'] <= self.radius*1000]
            
            # Remove blocks corresponding to schools, monitors, etc.
            blksinrange_df = blksinrange_df.loc[
                (~blksinrange_df['blkid'].str.contains('S')) &
                (~blksinrange_df['blkid'].str.contains('M'))]

            blksinrange_df['bkgrp'] = blksinrange_df['blkid'].astype(str).str[:12]
                        
            # Merge with ACS blockgroup data
            # Note: Not all blockgroups in blksinrange_gdf will be in the ACS blockgroup data
            commonACS_df = blksinrange_df.merge(
                self.acs_df.astype({'bkgrp': 'str'}), how='inner', left_on='bkgrp', right_on='bkgrp')

            # Identify any census blockgroups that are not in the ACS blockgroup data
            missing_df = blksinrange_df[(~blksinrange_df.bkgrp.isin(commonACS_df.bkgrp))].copy()
            
            if len(missing_df) == 0:
                acsinrange_df = commonACS_df
                
            else:
                # Add these missing blockgroups to the missing set
                missbkgrp = missing_df['bkgrp'].tolist()
                self.missingbkgrps.update(missbkgrp)
                
                # First try to default missing blockgroups to tracts
                missing_df['tract'] = missing_df['bkgrp'].str[:11]
                missing_w_tract = missing_df.merge(
                    self.acsCountyTract_df, how='inner', left_on='tract', right_on='ID')
                
                # Next, consider counties
                if (len(commonACS_df) + len(missing_w_tract)) != len(blksinrange_df):
                    missing_df['county'] = missing_df['bkgrp'].str[:5]
                    stillmissing_df = missing_df[(~missing_df.tract.isin(self.acsCountyTract_df.ID))]
                    missing_w_county = stillmissing_df.merge(
                        self.acsCountyTract_df, how='inner', left_on='county', right_on='ID')
                
                    if (len(commonACS_df) + len(missing_w_tract) + len(missing_w_county)) != len(blksinrange_df):
                        completelymissing_df = stillmissing_df[(~stillmissing_df.county.isin(self.acsCountyTract_df.ID))]
                        # messagebox.showinfo("Warning", "There are some census blocks that could not be matched to " +
                        #                     "ACS blockgroup or ACS default data.")
                    acsinrange_df = commonACS_df.append([missing_w_tract,missing_w_county], ignore_index=True)
                else:
                    acsinrange_df = commonACS_df.append(missing_w_tract, ignore_index=True)


            acs_columns = ['blkid', 'population', 'totalpop', 'p_minority', 'pnh_white', 'pnh_afr_am',
                           'pnh_am_ind', 'pnh_othmix', 'pt_hisp', 'p_agelt18', 'p_agegt64',
                           'p_2xpov', 'p_pov', 'age_25up', 'p_edulths', 'p_lingiso',
                           'age_univ', 'pov_univ', 'edu_univ', 'iso_univ', 'pov_fl', 'iso_fl']
            acsinrange_df = pd.DataFrame(acsinrange_df, columns=acs_columns)

                        
            #------------------------------------------------------------------------------------------
            # Create facility bin and tabulate population weighted demographic stats for each sub group.
            #------------------------------------------------------------------------------------------
            self.facility_bin = [[0]*16 for _ in range(2)]
            
            acsinrange_df['white'] = acsinrange_df['pnh_white'] * acsinrange_df['population']
            acsinrange_df['black'] = acsinrange_df['pnh_afr_am'] * acsinrange_df['population']
            acsinrange_df['amerind'] = acsinrange_df['pnh_am_ind'] * acsinrange_df['population']
            acsinrange_df['other'] = acsinrange_df['pnh_othmix'] * acsinrange_df['population']
            acsinrange_df['hisp'] = acsinrange_df['pt_hisp'] * acsinrange_df['population']
            acsinrange_df['agelt18'] = acsinrange_df['p_agelt18'] * acsinrange_df['population']
            acsinrange_df['agegt64'] = acsinrange_df['p_agegt64'] * acsinrange_df['population']
            acsinrange_df['age18to64'] = (100 - acsinrange_df['p_agelt18'] - acsinrange_df['p_agegt64']) * acsinrange_df['population']
            acsinrange_df['eduuniv'] = (acsinrange_df['edu_univ'] / acsinrange_df['totalpop']) * acsinrange_df['population']
            acsinrange_df['eduuniv100'] = (acsinrange_df['edu_univ'] / acsinrange_df['totalpop']) * acsinrange_df['population'] * 100
            acsinrange_df['povuniv100'] = (acsinrange_df['pov_univ'] / acsinrange_df['totalpop']) * acsinrange_df['population'] * 100 
            acsinrange_df['nohs'] = acsinrange_df['p_edulths'] * (acsinrange_df['edu_univ'] / acsinrange_df['totalpop']) \
                                                               * acsinrange_df['population']
            acsinrange_df['pov'] = acsinrange_df['p_pov'] * (acsinrange_df['pov_univ'] / acsinrange_df['totalpop']) \
                                                          * acsinrange_df['population']
            acsinrange_df['pov2x'] = acsinrange_df['p_2xpov'] * (acsinrange_df['pov_univ'] / acsinrange_df['totalpop']) \
                                                              * acsinrange_df['population']
            acsinrange_df['lingiso'] = acsinrange_df['p_lingiso'] * acsinrange_df['population']
            acsinrange_df['minority'] = acsinrange_df['p_minority'] * acsinrange_df['population']

            self.facility_bin[0][0] = acsinrange_df['population'].sum()
            self.facility_bin[0][1] = acsinrange_df[acsinrange_df['p_minority'].notna()]['population'].sum()
            self.facility_bin[0][2] = acsinrange_df[acsinrange_df['pnh_afr_am'].notna()]['population'].sum()
            self.facility_bin[0][3] = acsinrange_df[acsinrange_df['amerind'].notna()]['population'].sum()
            self.facility_bin[0][4] = acsinrange_df[acsinrange_df['pnh_othmix'].notna()]['population'].sum()
            self.facility_bin[0][5] = acsinrange_df[acsinrange_df['pt_hisp'].notna()]['population'].sum()
            self.facility_bin[0][6] = acsinrange_df[acsinrange_df['p_agelt18'].notna()]['population'].sum()
            self.facility_bin[0][7] = acsinrange_df[(acsinrange_df['p_agelt18'].notna()) &
                                                  (acsinrange_df['p_agegt64'].notna())]['population'].sum()
            self.facility_bin[0][8] = acsinrange_df[acsinrange_df['p_agegt64'].notna()]['population'].sum()
            self.facility_bin[0][9] = acsinrange_df[acsinrange_df['edu_univ'].notna()]['population'].sum()
            self.facility_bin[0][10] = acsinrange_df[(acsinrange_df['edu_univ'].notna()) &
                                       (acsinrange_df['p_edulths'].notna())]['eduuniv'].sum()
            self.facility_bin[0][11] = acsinrange_df[acsinrange_df['pov_univ'].notna()]['population'].sum()
            self.facility_bin[0][12] = acsinrange_df[(acsinrange_df['pov_univ'].notna()) &
                                       (acsinrange_df['p_2xpov'].notna())]['pov_univ'].sum()
            self.facility_bin[0][13] = acsinrange_df[acsinrange_df['p_lingiso'].notna()]['population'].sum()
            self.facility_bin[0][14] = acsinrange_df[acsinrange_df['p_minority'].notna()]['population'].sum()
            self.facility_bin[0][15] = acsinrange_df[acsinrange_df['pov_univ'].notna()]['population'].sum()
            
            self.facility_bin[1][1] = acsinrange_df[acsinrange_df['white'].notna()]['white'].sum()
            self.facility_bin[1][2] = acsinrange_df[acsinrange_df['black'].notna()]['black'].sum()
            self.facility_bin[1][3] = acsinrange_df[acsinrange_df['amerind'].notna()]['amerind'].sum()
            self.facility_bin[1][4] = acsinrange_df[acsinrange_df['other'].notna()]['other'].sum()
            self.facility_bin[1][5] = acsinrange_df[acsinrange_df['hisp'].notna()]['hisp'].sum()
            self.facility_bin[1][6] = acsinrange_df[acsinrange_df['agelt18'].notna()]['agelt18'].sum()
            self.facility_bin[1][7] = acsinrange_df[acsinrange_df['age18to64'].notna()]['age18to64'].sum()
            self.facility_bin[1][8] = acsinrange_df[acsinrange_df['agegt64'].notna()]['agegt64'].sum()  
            self.facility_bin[1][9] = acsinrange_df[acsinrange_df['eduuniv100'].notna()]['eduuniv100'].sum()
            self.facility_bin[1][10] = acsinrange_df[acsinrange_df['nohs'].notna()]['nohs'].sum()
            self.facility_bin[1][11] = acsinrange_df[acsinrange_df['pov'].notna()]['pov'].sum()
            self.facility_bin[1][12] = acsinrange_df[acsinrange_df['pov2x'].notna()]['pov2x'].sum()
            self.facility_bin[1][13] = acsinrange_df[acsinrange_df['lingiso'].notna()]['lingiso'].sum()
            self.facility_bin[1][14] = acsinrange_df[acsinrange_df['minority'].notna()]['minority'].sum()
            self.facility_bin[1][15] = acsinrange_df[acsinrange_df['povuniv100'].notna()]['povuniv100'].sum()            
                                    
            # Calculate facility averages by dividing population for each sub group
            for col_index in range(1, 16):
                if (self.facility_bin[0][col_index]) == 0:
                    self.facility_bin[1][col_index] = 0
                else:
                    self.facility_bin[1][col_index] = self.facility_bin[1][col_index] / (100 * self.facility_bin[0][col_index])
                    
            # Compute people counts
            self.facility_bin[0][15] = self.facility_bin[0][0] * self.facility_bin[1][15]
            for col_index in range(1, 15):
                # self.facility_bin[0][col_index] = self.facility_bin[0][0] * self.facility_bin[1][col_index]
                if col_index == 10:
                    self.facility_bin[0][col_index] = self.facility_bin[0][10] * self.facility_bin[1][col_index]
                else:
                    self.facility_bin[0][col_index] = self.facility_bin[0][0] * self.facility_bin[1][col_index]
        
            self.facility_bin[1][0] = ""

            # Write to facility sheet
            start_row = self.append_aggregated_data(
                self.facility_bin, self.worksheet_facility, self.formats, start_row)
            
            # Write facility to sortable sheet
            sort_bin = self.facility_bin[1]
            sort_bin[0] = self.facility_bin[0][0]
            col_idx = np.array(self.active_columns)
            slice = np.array(sort_bin)[col_idx]
            
            for col_num, data in enumerate(slice):
                format = self.formats['percentage'] if data <= 1 else self.formats['number']
                self.worksheet_sort.write_number(sort_row, col_num+3, data, format)
            sort_row = sort_row + 1

            # Add facility data to run group bin
            if self.rungroup_bin == None:
                self.rungroup_bin = [[0]*16 for _ in range(2)]
            self.tabulate_rungroup_data(acsinrange_df)

            # Put blkid's from acsinrange_df into unique list of used blocks for later use by rungroup
            acsblk_list = acsinrange_df['blkid'].tolist()
            allblks = self.used_blocks
            allblks.extend(acsblk_list)
            self.used_blocks = list(set(allblks))


        #----------- Process the run group bin --------------------
        
        # Calculate averages by dividing population for each sub group
        for col_index in range(1, 16):
            if (self.rungroup_bin[0][col_index]) == 0:
                self.rungroup_bin[1][col_index] = 0
            else:
                self.rungroup_bin[1][col_index] = self.rungroup_bin[1][col_index] / (100 * self.rungroup_bin[0][col_index])

        # Compute people counts
        self.rungroup_bin[0][15] = self.rungroup_bin[0][0] * self.rungroup_bin[1][15]
        for col_index in range(1, 15):
            # self.rungroup_bin[0][col_index] = self.rungroup_bin[0][0] * self.rungroup_bin[1][col_index]
            if col_index == 10:
                self.rungroup_bin[0][col_index] = self.rungroup_bin[0][10] * self.rungroup_bin[1][col_index]
            else:
                self.rungroup_bin[0][col_index] = self.rungroup_bin[0][0] * self.rungroup_bin[1][col_index]

        self.rungroup_bin[1][0] = ""

        # Write to facility sheet
        self.worksheet_facility.write_string(start_row+1, 0, 'Run group total', self.formats['sub_header_3'])
        start_row = self.append_aggregated_data(
            self.rungroup_bin, self.worksheet_facility, self.formats, start_row+1)

        # Write to sortable sheet
        self.worksheet_sort.write_string(sort_row, 0, 'Run group total', self.formats['sub_header_3'])
        sort_bin = self.rungroup_bin[1]
        sort_bin[0] = self.rungroup_bin[0][0]
        col_idx = np.array(self.active_columns)
        slice = np.array(sort_bin)[col_idx]

        for col_num, data in enumerate(slice):
            format = self.formats['percentage'] if data <= 1 else self.formats['number']
            self.worksheet_sort.write_number(sort_row, col_num+3, data, format)
        sort_row = sort_row + 1
        
    # Create Workbook
    # Final workbook should have similar formatting as ej tables, with two rows for nationwide
    # demographics (population and percentages) and two rows for each facility provided in the
    # original faclist. 
    def create_workbook(self):
        output_dir = self.fullpath
        if not (os.path.exists(output_dir) or os.path.isdir(output_dir)):
            os.mkdir(output_dir)
        filename = os.path.join(output_dir, self.filename_entry + '.xlsx')
        self.workbook = xlsxwriter.Workbook(filename)
        self.worksheet_facility = self.workbook.add_worksheet('Facility Demographics')
        self.worksheet_sort = self.workbook.add_worksheet('Sortable %')
        self.formats = self.create_formats(self.workbook)

        #------------ Facility Spreadsheet ----------------------------------------------

        tablename = 'Population Demographics within ' + str(self.radius) + ' km of Source Facilities'
        
        column_headers = ['Total Population', 'White', 'Minority', 'African American',
                          'Native American', 'Other and Multiracial', 'Hispanic or Latino',
                          'Age (Years)\n0-17', 'Age (Years)\n18-64', 'Age (Years)\n>=65',
                          'People Living Below the Poverty Level', 'Total Number >= 25 Years Old',
                          'Number >= 25 Years Old without a High School Diploma',
                          'People Living in Linguistic Isolation']

        firstcol = 'A'
        lastcol = chr(ord(firstcol) + len(column_headers))
        top_header_coords = firstcol+'1:'+lastcol+'1'

        # Increase the cell size of the merged cells to highlight the formatting.
        self.worksheet_facility.set_column(top_header_coords, 12)
        self.worksheet_facility.set_row(0, 30)

        # Create top level header
        self.worksheet_facility.merge_range(top_header_coords, tablename, self.formats['top_header'])

        # Create column headers
        self.worksheet_facility.merge_range("A2:A3", 'Population Basis', self.formats['sub_header_2'])
        self.worksheet_facility.merge_range("A4:A5", 'Nationwide', self.formats['sub_header_3'])
        self.worksheet_facility.merge_range("B2:N2", 'Demographic Group',  self.formats['sub_header_3'])

        self.worksheet_facility.set_row(2, 72, self.formats['sub_header_2'])
        for col_num, data in enumerate(column_headers):
            self.worksheet_facility.write(2, col_num+1, data)

        # Add Facility Names
        facname_list = self.faclist_df['facility_id'].tolist()
        row_num = 6
        for index, data in enumerate(facname_list):
            self.worksheet_facility.merge_range(row_num, 0, row_num + 1, 0, data, self.formats['sub_header_3'])
            row_num = row_num + 2

        last_data_row = 2 * len(facname_list) + 10

        # Create notes
        first_notes_row = last_data_row + 1
        # last_notes_row = first_notes_row + 4
        firstcol = 'A'
        lastcol = chr(ord(firstcol) + len(column_headers))
        notes_coords = firstcol+str(first_notes_row)+':'+lastcol+str(first_notes_row)
        # notes_coords = firstcol+str(first_notes_row)+':'+lastcol+str(last_notes_row)
        self.worksheet_facility.merge_range(notes_coords, 'Notes:\n' + \
          '* Total nationwide population includes all 50 states plus Puerto Rico.\n' + \
          '* Distributions by race, ethnicity, age, education, income and linguistic isolation are based on ' + \
          "demographic information at the census block group level, provided by the Census' American Community Survey (ACS) 5-year averages. Demographic percentages based on different averages may differ.\n" + \
          '* The minority population includes people identifying as African American, Native American, Other ' + \
          'and Multiracial, or Hispanic/Latino. Measures are taken to avoid double counting of people identifying ' + \
          'as both Hispanic/Latino and a racial minority. ' + \
          'In order to avoid double counting, the "Hispanic or Latino" category is treated as a distinct ' + \
          'demographic category for these analyses. A person is identified as one of five racial/ethnic ' + \
          'categories above: White, African American, Native American, Other and Multiracial, or Hispanic/Latino.\n' \
          ,  self.formats['notes'])

        self.worksheet_facility.set_row(first_notes_row-1, 120)


        #------------ Sortable Spreadsheet ----------------------------------------------

        sort_headers = ['Facility ID', 'Longitude', 'Latitude', 'Total Population', 'White', 
                          'Minority', 'African American',
                          'Native American', 'Other and Multiracial', 'Hispanic or Latino',
                          'Age (Years)\n0-17', 'Age (Years)\n18-64', 'Age (Years)\n>=65',
                          'People Living Below the Poverty Level', 'Total Number >= 25 Years Old',
                          'Number >= 25 Years Old without a High School Diploma',
                          'People Living in Linguistic Isolation']
        
        firstcol = 'A'
        lastcol = chr(ord(firstcol) + len(sort_headers))
        top_header_coords = firstcol+'1:'+lastcol+'1'

        # Increase the column width.
        self.worksheet_sort.set_column(top_header_coords, 12)
              
        # Create column headers
        self.worksheet_sort.set_row(0, 72, self.formats['sub_header_2'])
        for col_num, data in enumerate(sort_headers):
            self.worksheet_sort.write(0, col_num, data)
        
        # Add Facility ID, Lat, Lon
        facname_list = self.faclist_df['facility_id'].tolist()
        row_num = 1
        for index, row in self.faclist_df.iterrows():
            self.worksheet_sort.write_string(row_num, 0, row['facility_id'], self.formats['sub_header_3'])
            self.worksheet_sort.write_number(row_num, 1, row['lon'], self.formats['sub_header_3'])
            self.worksheet_sort.write_number(row_num, 2, row['lat'], self.formats['sub_header_3'])
            row_num = row_num + 1
                
        
    def close_workbook(self):
        self.workbook.close()

