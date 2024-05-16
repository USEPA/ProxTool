# -*- coding: utf-8 -*-
"""
Created on Wed Jul 21 14:06:32 2021

@author: CCook
"""
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
import os
import csv
import numpy as np
import xlsxwriter
import pandas as pd

from copy import deepcopy
from decimal import ROUND_HALF_UP, Decimal, getcontext
from com.sca.ca.support.UTM import *


# Describe Demographics Within Range of Specified Facilities
class FacilityProximityAssessment:

    def __init__(self, filename_entry, output_dir, faclist_df, radius, census_df, acs_df, 
                 acsDefault_df):

        # Output path
        self.filename_entry = str(filename_entry) + '.xlsx'
        self.fullpath = output_dir
        self.faclist_df = faclist_df
        self.censusblks_df = census_df
        self.acs_df = acs_df
        self.acsDefault_df = acsDefault_df
        self.formats = None
        self.facility_bin = None
        self.national_bin = None
        self.rungroup_bin = None
        
        # Initialize list of used blocks
        self.used_blocks = []
        
        # Initialize missing blockgroups list
        self.missingbkgrps = []

        # Specify range in km
        self.radius = int(radius)

        # Identify the relevant column indexes from the national and facility bins
        self.active_columns = [0, 1, 15, 2, 3, 4, 5, 6, 7, 8, 9, 12, 13, 10, 11, 14]
        
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

        formats['sub_header_5'] = workbook.add_format({
            'bold': 0,
            'align': 'center',
            'valign': 'bottom',
            'text_wrap': 1})

        formats['sub_header_6'] = workbook.add_format({
            'bold': 0,
            'align': 'left',
            'valign': 'vcenter',
            'text_wrap': 1})

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
        
        formats['superscript'] = workbook.add_format({'font_script': 1})

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
                    # Override format for the National row
                    if numrows==1:
                        format = formats['percentage'] if col > 0 else formats['number'] 
                    # Override format for the rungroup and facility total population percentage
                    if row==1 and col==0:
                        format = formats['int_percentage']
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
        self.rungroup_bin[0][4] += rungroup_df[rungroup_df['asian'].notna()]['population'].sum()
        self.rungroup_bin[0][5] += rungroup_df[rungroup_df['pnh_othmix'].notna()]['population'].sum()
        self.rungroup_bin[0][6] += rungroup_df[rungroup_df['pt_hisp'].notna()]['population'].sum()
        self.rungroup_bin[0][7] += rungroup_df[rungroup_df['p_agelt18'].notna()]['population'].sum()
        self.rungroup_bin[0][8] += rungroup_df[(rungroup_df['p_agelt18'].notna()) &
                                              (rungroup_df['p_agegt64'].notna())]['population'].sum()
        self.rungroup_bin[0][9] += rungroup_df[rungroup_df['p_agegt64'].notna()]['population'].sum()
        self.rungroup_bin[0][10] += rungroup_df[rungroup_df['edu_univ'].notna()]['population'].sum()
        self.rungroup_bin[0][11] += rungroup_df[(rungroup_df['edu_univ'].notna()) &
                                   (rungroup_df['p_edulths'].notna())]['eduuniv'].sum()
        self.rungroup_bin[0][12] += rungroup_df[rungroup_df['pov_univ'].notna()]['population'].sum()
        self.rungroup_bin[0][13] += rungroup_df[(rungroup_df['pov_univ'].notna()) &
                                   (rungroup_df['p_2xpov'].notna())]['population'].sum()
        self.rungroup_bin[0][14] += rungroup_df[rungroup_df['p_lingiso'].notna()]['population'].sum()
        self.rungroup_bin[0][15] += rungroup_df[rungroup_df['p_minority'].notna()]['population'].sum()
        self.rungroup_bin[0][16] += rungroup_df[rungroup_df['pov_univ'].notna()]['population'].sum()
        
        self.rungroup_bin[1][1] += rungroup_df[rungroup_df['white'].notna()]['white'].sum()
        self.rungroup_bin[1][2] += rungroup_df[rungroup_df['black'].notna()]['black'].sum()
        self.rungroup_bin[1][3] += rungroup_df[rungroup_df['amerind'].notna()]['amerind'].sum()
        self.rungroup_bin[1][4] += rungroup_df[rungroup_df['asian'].notna()]['asian'].sum()
        self.rungroup_bin[1][5] += rungroup_df[rungroup_df['other'].notna()]['other'].sum()
        self.rungroup_bin[1][6] += rungroup_df[rungroup_df['hisp'].notna()]['hisp'].sum()
        self.rungroup_bin[1][7] += rungroup_df[rungroup_df['agelt18'].notna()]['agelt18'].sum()
        self.rungroup_bin[1][8] += rungroup_df[rungroup_df['age18to64'].notna()]['age18to64'].sum()
        self.rungroup_bin[1][9] += rungroup_df[rungroup_df['agegt64'].notna()]['agegt64'].sum()  
        self.rungroup_bin[1][10] += rungroup_df[rungroup_df['eduuniv100'].notna()]['eduuniv100'].sum()
        self.rungroup_bin[1][11] += rungroup_df[rungroup_df['nohs'].notna()]['nohs'].sum()
        self.rungroup_bin[1][12] += rungroup_df[rungroup_df['pov'].notna()]['pov'].sum()
        self.rungroup_bin[1][13] += rungroup_df[rungroup_df['pov2x'].notna()]['pov2x'].sum()
        self.rungroup_bin[1][14] += rungroup_df[rungroup_df['lingiso'].notna()]['lingiso'].sum()
        self.rungroup_bin[1][15] += rungroup_df[rungroup_df['minority'].notna()]['minority'].sum()
        self.rungroup_bin[1][16] += rungroup_df[rungroup_df['povuniv100'].notna()]['povuniv100'].sum()            
        



    def create(self):
        self.create_workbook()
        self.calculate_distances()
        self.close_workbook()
                
        # Write out any defaulted blockgroups
        if len(self.missingbkgrps) > 0:
            missfname = os.path.splitext(self.filename_entry)[0] + '_' + 'defaulted_block_groups' + '_' +\
                        str(self.radius) + 'km.txt'
            misspath = os.path.join(self.fullpath, missfname)

            
            with open(misspath, 'w') as f:
                wr = csv.writer(f, delimiter="-")
                wr.writerows(self.missingbkgrps)
        

    def calculate_distances(self):

        # Initialize starting data rows for the facility and sortable sheets (zero-indexed)
        start_row = 2
        sort_row = 3

        #------------------------------------------------------------------------------------------
        # Create national bin and tabulate population weighted demographic stats for each sub group.
        #------------------------------------------------------------------------------------------
        self.national_bin = [[0]*17 for _ in range(2)]

        national_acs = self.acs_df
                
        national_acs['white'] = national_acs['pnh_white'] * national_acs['totalpop']
        national_acs['black'] = national_acs['pnh_afr_am'] * national_acs['totalpop']
        national_acs['amerind'] = national_acs['pnh_am_ind'] * national_acs['totalpop']
        national_acs['asian'] = national_acs['pnh_asian'] * national_acs['totalpop']
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
        self.national_bin[0][4] = national_acs[national_acs['asian'].notna()]['totalpop'].sum()
        self.national_bin[0][5] = national_acs[national_acs['pnh_othmix'].notna()]['totalpop'].sum()
        self.national_bin[0][6] = national_acs[national_acs['pt_hisp'].notna()]['totalpop'].sum()
        self.national_bin[0][7] = national_acs[national_acs['p_agelt18'].notna()]['totalpop'].sum()
        self.national_bin[0][8] = national_acs[(national_acs['p_agelt18'].notna()) &
                                              (national_acs['p_agegt64'].notna())]['totalpop'].sum()
        self.national_bin[0][9] = national_acs[national_acs['p_agegt64'].notna()]['totalpop'].sum()
        self.national_bin[0][10] = national_acs[national_acs['edu_univ'].notna()]['totalpop'].sum()
        self.national_bin[0][11] = national_acs[(national_acs['edu_univ'].notna()) &
                                   (national_acs['p_edulths'].notna())]['edu_univ'].sum()
        self.national_bin[0][12] = national_acs[national_acs['pov_univ'].notna()]['totalpop'].sum()
        self.national_bin[0][13] = national_acs[(national_acs['pov_univ'].notna()) &
                                   (national_acs['p_2xpov'].notna())]['totalpop'].sum()
        self.national_bin[0][14] = national_acs[national_acs['p_lingiso'].notna()]['totalpop'].sum()
        self.national_bin[0][15] = national_acs[national_acs['p_minority'].notna()]['totalpop'].sum()
        self.national_bin[0][16] = national_acs[national_acs['pov_univ'].notna()]['totalpop'].sum()

        self.national_bin[1][1] = national_acs[national_acs['white'].notna()]['white'].sum()
        self.national_bin[1][2] = national_acs[national_acs['black'].notna()]['black'].sum()
        self.national_bin[1][3] = national_acs[national_acs['amerind'].notna()]['amerind'].sum()
        self.national_bin[1][4] = national_acs[national_acs['asian'].notna()]['asian'].sum()
        self.national_bin[1][5] = national_acs[national_acs['other'].notna()]['other'].sum()
        self.national_bin[1][6] = national_acs[national_acs['hisp'].notna()]['hisp'].sum()
        self.national_bin[1][7] = national_acs[national_acs['agelt18'].notna()]['agelt18'].sum()
        self.national_bin[1][8] = national_acs[national_acs['age18to64'].notna()]['age18to64'].sum()
        self.national_bin[1][9] = national_acs[national_acs['agegt64'].notna()]['agegt64'].sum()  
        self.national_bin[1][10] = national_acs[national_acs['eduuniv100'].notna()]['eduuniv100'].sum()
        self.national_bin[1][11] = national_acs[national_acs['nohs'].notna()]['nohs'].sum()
        self.national_bin[1][12] = national_acs[national_acs['pov'].notna()]['pov'].sum()
        self.national_bin[1][13] = national_acs[national_acs['pov2x'].notna()]['pov2x'].sum()
        self.national_bin[1][14] = national_acs[national_acs['lingiso'].notna()]['lingiso'].sum()
        self.national_bin[1][15] = national_acs[national_acs['minority'].notna()]['minority'].sum()
        self.national_bin[1][16] = national_acs[national_acs['povuniv100'].notna()]['povuniv100'].sum()
 
        # Only demographic percentages of the National bin will be reported

        # Calculate fractions by dividing population for each sub group.
        # Note that pov and 2xpov are divided by pov_univ and noHS is divided by edu_univ.
        for index in range(1, 17):
            self.national_bin[0][index] = self.national_bin[1][index] / (100 * self.national_bin[0][index])
            # if index == 12:
            #     self.national_bin[0][index] = self.national_bin[1][index] / (100 * self.national_bin[0][0])
            # else:
            #     self.national_bin[0][index] = self.national_bin[1][index] / (100 * self.national_bin[0][index])
 
        # Delete index 1 from the Natinal bin list. Only keeping percentages.
        del self.national_bin[-1]
        
        
        # Write to facility sheet and leave rows for the run group total
        start_row = self.append_aggregated_data(
            self.national_bin, self.worksheet_facility, self.formats, start_row) + 5

        # Write to sortable sheet (row 1)
        data = deepcopy(self.national_bin)
        # Keep relevant columns
        newdata = [data[0][i] for i in self.active_columns]
        for col in range(0, len(newdata)):
            value = float(newdata[col])
            format = self.formats['percentage'] if col > 0 else self.formats['number'] 
            self.worksheet_sort.write_number(1, col+3, value, format)


        
        # Process each facility
        for index, row in self.faclist_df.iterrows():
            
            print('Calculating proximity for facility: ' + self.faclist_df['facility_id'][index])
                            
            self.facility_bin = [[0]*17 for _ in range(2)]
            
            fac_lat = row['lat']
            fac_lon = row['lon']
                        
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
                        
            # Remove blocks corresponding to schools, monitors, and user receptors.
            blksinrange_df = blksinrange_df.loc[
                (~blksinrange_df['blkid'].str.contains('S')) &
                (~blksinrange_df['blkid'].str.contains('M')) &
                (~blksinrange_df['blkid'].str.contains('U'))]

            blksinrange_df['bkgrp'] = blksinrange_df['blkid'].astype(str).str[:12]
                        
            # Merge with ACS blockgroup data
            # Note: Not all blockgroups in blksinrange_df will be in the ACS blockgroup data
            acsinrange_df = blksinrange_df.merge(
                self.acs_df.astype({'bkgrp': 'str'}), how='inner', left_on='bkgrp', right_on='bkgrp')

            # Identify any unique census blockgroups that are not in the ACS blockgroup data
            missing_df = blksinrange_df[(~blksinrange_df.bkgrp.isin(acsinrange_df.bkgrp))].copy()
            missing_list = missing_df['bkgrp'].unique().tolist()
                            
            if len(missing_df) > 0:
                
                # Look for the missing block groups in the default file using block group
                bkgrp_defaults_df = self.acsDefault_df[self.acsDefault_df['rectype'] == 'BLKGRP']
                found_df = bkgrp_defaults_df[bkgrp_defaults_df['tct'].isin(missing_list)]
                
                # Record the block groups defaulted by nearest block group
                for b in found_df['tct'].unique().tolist():
                    if b not in self.missingbkgrps:
                        self.missingbkgrps.append([b, 'Defaulted to nearest block group'])
                
                acsinrange_df = pd.concat([acsinrange_df, found_df], ignore_index=True)
                
                # Remove defaulted blockgroups from the missing list
                missing_list = [bg for bg in missing_list if bg not in found_df['tct'].to_list()]

                # If there are still missing block groups, then use tract defaults
                if len(missing_list) > 0:
                    tract_defaults_df = self.acsDefault_df[self.acsDefault_df['rectype'] == 'TCT']
                    found_df = tract_defaults_df[tract_defaults_df['tct'].isin(missing_list)]

                    # Record the block groups defaulted by tract
                    for b in found_df['tct'].unique().tolist():
                        if b not in self.missingbkgrps:
                            self.missingbkgrps.append([b, 'Defaulted by tract'])

                    # Remove defaulted blockgroups from the missing list
                    missing_list = [bg for bg in missing_list if bg not in found_df['tct'].to_list()]

                    # If there are still missing block groups, then use county defaults
                    if len(missing_list) > 0:
                        cty_defaults_df = self.acsDefault_df[self.acsDefault_df['rectype'] == 'CTY']
                        found_df = cty_defaults_df[cty_defaults_df['tct'].isin(missing_list)]
    
                        # Record the block groups defaulted by county
                        for b in found_df['tct'].unique().tolist():
                            if b not in self.missingbkgrps:
                                self.missingbkgrps.append([b, 'Defaulted by county'])
    
                        # Remove defaulted blockgroups from the missing list
                        missing_list = [bg for bg in missing_list if bg not in found_df['tct'].to_list()]

                        # If there are still missing block groups, then this is an error
                        if len(missing_list) > 0:
                            for b in missing_list:
                                self.missingbkgrps.append([b, 'Error! Could not be defaulted'])
                    

            # Set column names
            acs_columns = ['blkid', 'population', 'totalpop', 'p_minority', 'pnh_white', 'pnh_afr_am',
                           'pnh_am_ind', 'pnh_asian', 'pt_hisp', 'pnh_othmix', 'p_agelt18', 'p_agegt64',
                           'p_2xpov', 'p_pov', 'p_edulths', 'p_lingiso',
                           'pov_univ', 'edu_univ', 'iso_univ']
            acsinrange_df = pd.DataFrame(acsinrange_df, columns=acs_columns)
                        
            #------------------------------------------------------------------------------------------
            # Create facility bin and tabulate population weighted demographic stats for each sub group.
            #------------------------------------------------------------------------------------------
            self.facility_bin = [[0]*17 for _ in range(2)]
            
            acsinrange_df['white'] = acsinrange_df['pnh_white'] * acsinrange_df['population']
            acsinrange_df['black'] = acsinrange_df['pnh_afr_am'] * acsinrange_df['population']
            acsinrange_df['amerind'] = acsinrange_df['pnh_am_ind'] * acsinrange_df['population']
            acsinrange_df['asian'] = acsinrange_df['pnh_asian'] * acsinrange_df['population']
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
            self.facility_bin[0][4] = acsinrange_df[acsinrange_df['asian'].notna()]['population'].sum()
            self.facility_bin[0][5] = acsinrange_df[acsinrange_df['pnh_othmix'].notna()]['population'].sum()
            self.facility_bin[0][6] = acsinrange_df[acsinrange_df['pt_hisp'].notna()]['population'].sum()
            self.facility_bin[0][7] = acsinrange_df[acsinrange_df['p_agelt18'].notna()]['population'].sum()
            self.facility_bin[0][8] = acsinrange_df[(acsinrange_df['p_agelt18'].notna()) &
                                                  (acsinrange_df['p_agegt64'].notna())]['population'].sum()
            self.facility_bin[0][9] = acsinrange_df[acsinrange_df['p_agegt64'].notna()]['population'].sum()
            self.facility_bin[0][10] = acsinrange_df[acsinrange_df['edu_univ'].notna()]['population'].sum()
            self.facility_bin[0][11] = acsinrange_df[(acsinrange_df['edu_univ'].notna()) &
                                                     (acsinrange_df['p_edulths'].notna())]['eduuniv'].sum()
            self.facility_bin[0][12] = acsinrange_df[acsinrange_df['pov_univ'].notna()]['population'].sum()
            self.facility_bin[0][13] = acsinrange_df[(acsinrange_df['pov_univ'].notna()) &
                                                     (acsinrange_df['p_2xpov'].notna())]['population'].sum()
            self.facility_bin[0][14] = acsinrange_df[acsinrange_df['p_lingiso'].notna()]['population'].sum()
            self.facility_bin[0][15] = acsinrange_df[acsinrange_df['p_minority'].notna()]['population'].sum()
            self.facility_bin[0][16] = acsinrange_df[acsinrange_df['pov_univ'].notna()]['population'].sum()
            
            self.facility_bin[1][1] = acsinrange_df[acsinrange_df['white'].notna()]['white'].sum()
            self.facility_bin[1][2] = acsinrange_df[acsinrange_df['black'].notna()]['black'].sum()
            self.facility_bin[1][3] = acsinrange_df[acsinrange_df['amerind'].notna()]['amerind'].sum()
            self.facility_bin[1][4] = acsinrange_df[acsinrange_df['asian'].notna()]['asian'].sum()
            self.facility_bin[1][5] = acsinrange_df[acsinrange_df['other'].notna()]['other'].sum()
            self.facility_bin[1][6] = acsinrange_df[acsinrange_df['hisp'].notna()]['hisp'].sum()
            self.facility_bin[1][7] = acsinrange_df[acsinrange_df['agelt18'].notna()]['agelt18'].sum()
            self.facility_bin[1][8] = acsinrange_df[acsinrange_df['age18to64'].notna()]['age18to64'].sum()
            self.facility_bin[1][9] = acsinrange_df[acsinrange_df['agegt64'].notna()]['agegt64'].sum()  
            self.facility_bin[1][10] = acsinrange_df[acsinrange_df['eduuniv100'].notna()]['eduuniv100'].sum()
            self.facility_bin[1][11] = acsinrange_df[acsinrange_df['nohs'].notna()]['nohs'].sum()
            self.facility_bin[1][12] = acsinrange_df[acsinrange_df['pov'].notna()]['pov'].sum()
            self.facility_bin[1][13] = acsinrange_df[acsinrange_df['pov2x'].notna()]['pov2x'].sum()
            self.facility_bin[1][14] = acsinrange_df[acsinrange_df['lingiso'].notna()]['lingiso'].sum()
            self.facility_bin[1][15] = acsinrange_df[acsinrange_df['minority'].notna()]['minority'].sum()
            self.facility_bin[1][16] = acsinrange_df[acsinrange_df['povuniv100'].notna()]['povuniv100'].sum()            
                                    
            # Calculate facility averages by dividing population for each sub group
            for col_index in range(1, 17):
                if (self.facility_bin[0][col_index]) == 0:
                    self.facility_bin[1][col_index] = 0
                else:
                    self.facility_bin[1][col_index] = self.facility_bin[1][col_index] / (100 * self.facility_bin[0][col_index])
            # Hard code facility total population fraction as 1 (100%)
            self.facility_bin[1][0] = 1
                    
            # Compute people counts
            self.facility_bin[0][16] = self.facility_bin[0][0] * self.facility_bin[1][16]
            for col_index in range(1, 16):
                self.facility_bin[0][col_index] = self.facility_bin[0][col_index] * self.facility_bin[1][col_index]
                # if col_index == 11:
                #     self.facility_bin[0][col_index] = self.facility_bin[0][11] * self.facility_bin[1][col_index]
                # else:
                #     self.facility_bin[0][col_index] = self.facility_bin[0][0] * self.facility_bin[1][col_index]
        

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
                self.rungroup_bin = [[0]*17 for _ in range(2)]
            self.tabulate_rungroup_data(acsinrange_df)

            # Put blkid's from acsinrange_df into unique list of used blocks for later use by rungroup
            acsblk_list = acsinrange_df['blkid'].tolist()
            allblks = self.used_blocks
            allblks.extend(acsblk_list)
            self.used_blocks = list(set(allblks))


        #Temp - write used_blocks to a file
        with open('C:\\Git_CA\\ca\output\\PrimaryCopper_blocks.txt', 'w') as f:
            for line in self.used_blocks:
                f.write(f"{line}\n")        

        
        #----------- Process the run group bin --------------------
        
        # Calculate averages by dividing population for each sub group
        for col_index in range(1, 17):
            if (self.rungroup_bin[0][col_index]) == 0:
                self.rungroup_bin[1][col_index] = 0
            else:
                self.rungroup_bin[1][col_index] = self.rungroup_bin[1][col_index] / (100 * self.rungroup_bin[0][col_index])
        # Hard code run group total population rraction as 1 (100%)
        self.rungroup_bin[1][0] = 1

        # Compute people counts
        self.rungroup_bin[0][16] = self.rungroup_bin[0][0] * self.rungroup_bin[1][16]
        for col_index in range(1, 16):
            self.rungroup_bin[0][col_index] = self.rungroup_bin[0][col_index] * self.rungroup_bin[1][col_index]
            # if col_index == 11:
            #     self.rungroup_bin[0][col_index] = self.rungroup_bin[0][11] * self.rungroup_bin[1][col_index]
            # else:
            #     self.rungroup_bin[0][col_index] = self.rungroup_bin[0][0] * self.rungroup_bin[1][col_index]


        #------- Write to facility sheet --------------
        self.worksheet_facility.write_rich_string("A6", 'Run group total (pop.)',  self.formats['superscript']
                                                  , ' 8', self.formats['sub_header_6'])
        self.worksheet_facility.write_rich_string("A7", 'Run group total (pop.%)',  self.formats['superscript']
                                                  , ' 8', self.formats['sub_header_6'])
        start_row = self.append_aggregated_data(
            self.rungroup_bin, self.worksheet_facility, self.formats, 5)

        # Write to sortable sheet
        self.worksheet_sort.write_string(1, 0, 'Nationwide Demographics (2018-2022 ACS)', self.formats['sub_header_6'])
        self.worksheet_sort.write_string(2, 0, 'Run group total', self.formats['sub_header_6'])
        self.worksheet_sort.write_string(1, 1, 'N/A', self.formats['sub_header_5'])
        self.worksheet_sort.write_string(2, 1, 'N/A', self.formats['sub_header_5'])
        self.worksheet_sort.write_string(1, 2, 'N/A', self.formats['sub_header_5'])
        self.worksheet_sort.write_string(2, 2, 'N/A', self.formats['sub_header_5'])
        sort_bin = self.rungroup_bin[1]
        sort_bin[0] = self.rungroup_bin[0][0]
        col_idx = np.array(self.active_columns)
        slice = np.array(sort_bin)[col_idx]

        for col_num, data in enumerate(slice):
            format = self.formats['percentage'] if data <= 1 else self.formats['number']
            self.worksheet_sort.write_number(2, col_num+3, data, format)
        # sort_row = sort_row + 1
        
    # Create Workbook
    # Final workbook should have similar formatting as ej tables, with two rows for nationwide
    # demographics (population and percentages) and two rows for each facility provided in the
    # original faclist. 
    def create_workbook(self):
        output_dir = self.fullpath
        if not (os.path.exists(output_dir) or os.path.isdir(output_dir)):
            os.mkdir(output_dir)
        filename = os.path.join(output_dir, self.filename_entry)
        self.workbook = xlsxwriter.Workbook(filename)
        self.worksheet_readme = self.workbook.add_worksheet('Background ReadMe')
        self.worksheet_facility = self.workbook.add_worksheet('Facility Demographics')
        self.worksheet_sort = self.workbook.add_worksheet('Sortable %')
        self.formats = self.create_formats(self.workbook)

        #------------ Facility Spreadsheet ----------------------------------------------

        tablename = 'Population Demographics within ' + str(self.radius) + ' km of Source Facilities \u00B9'
        
        column_headers = ['Total Population', 'White', 'People of Color', 'Black',
                          'American Indian or Alaska Native', 'Asian', 'Other and Multiracial', 'Hispanic or Latino',
                          'Age (Years)\n0-17', 'Age (Years)\n18-64', 'Age (Years)\n>=65',
                          'People Living Below the Poverty Level', 
                          'People Living Below Twice the Poverty Level',
                          'People >= 25 Years Old',
                          'People >= 25 Years Old without a High School Diploma',
                          'People Living in Limited English Speaking Households']

        firstcol = 'A'
        lastcol = chr(ord(firstcol) + len(column_headers))
        top_header_coords = firstcol+'1:'+lastcol+'1'

        # # Add static content to the readme tab        
        self.worksheet_readme.write_string("A2", 'BACKGROUND:', self.formats['sub_header_4'])
        self.worksheet_readme.write_string("A4", "This analysis used the Proximity Tool to (1) identify all census blocks within a")
        self.worksheet_readme.write_string("A5", "specified radius of the latitude/longitude location of each facility, and then (2) link each block")
        self.worksheet_readme.write_string("A6", "with census-based demographic data. In addition to facility-specific demographics, the Proximity ")
        self.worksheet_readme.write_string("A7", "Tool also computes the demographic composition of the population within the specified radius for all ")
        self.worksheet_readme.write_string("A8", "facilities in the run group as a whole (e.g., source category-wide). Finally, this analysis allows for comparison of ")
        self.worksheet_readme.write_string("A9", "the facility-specific and source category-wide demographics at the specified radius to the nationwide demographics ")
        self.worksheet_readme.write_string("A10", "of the U.S. population. The Proximity Tool was created by SC&A Inc. in 2021 under contract to the U.S. EPA ")
        self.worksheet_readme.write_string("A11", "and has been updated most recently based on the 2020 Decennial Census population (at the census block level) ")
        self.worksheet_readme.write_string("A12", "and the 2018-2022 American Community Survey demographics (at the census block group level).")

        # self.worksheet_readme.merge_range("A4:I13", background_text, self.formats['notes'])

    # Set first column width to 26; all others to 12
        self.worksheet_facility.set_column("A1:A1", 26)
        self.worksheet_facility.set_column("B1:"+lastcol+"1", 12)
        
        # Increase the cell size of the top row to highlight the formatting.
        self.worksheet_facility.set_row(0, 30)

        # Create top level header
        self.worksheet_facility.merge_range(top_header_coords, tablename, self.formats['top_header'])

        # Create column headers
        self.worksheet_facility.write_string("A2", 'Population Basis', self.formats['sub_header_2'])
        self.worksheet_facility.write_string("A3", 'Nationwide Demographics (2018-2022 ACS)', self.formats['sub_header_6'])
        self.worksheet_facility.write_rich_string("A4", 'Nationwide (2020 Decennial Census)',  self.formats['superscript']
                                                  , ' 7', self.formats['sub_header_6'])
        self.worksheet_facility.write_number("B4", 334753155, self.formats['number'])
        # self.worksheet_facility.merge_range("B2:N2", '',  self.formats['sub_header_3'])
        # self.worksheet_facility.write_rich_string("B2", 'Demographic Group',  self.formats['superscript']
        #                                           , '1', self.formats['sub_header_3'])

        self.worksheet_facility.set_row(1, 78, self.formats['sub_header_2'])
        # define superscripts of demographic headers
        ss = {2:'2', 7:'3', 11:'4', 12:'4', 14:'5', 15:'6'}
        for col_num, data in enumerate(column_headers):
            if col_num in ss:
                # headers with superscripts
                self.worksheet_facility.write_rich_string(1, col_num+1, data,  self.formats['superscript']
                                                         , ' '+ss[col_num], self.formats['sub_header_2'])
            else:   
                self.worksheet_facility.write(1, col_num+1, data)

        # Add Facility Names
        facname_list = self.faclist_df['facility_id'].tolist()
        row_num = 8
        for index, data in enumerate(facname_list):
            self.worksheet_facility.write_string(row_num, 0, data + ' (pop.)', self.formats['sub_header_6'])
            row_num = row_num + 1
            self.worksheet_facility.write_string(row_num, 0, data + ' (pop.%)', self.formats['sub_header_6'])
            row_num = row_num + 1

        last_data_row = 2 * len(facname_list) + 9

        # Create notes
        first_notes_row = last_data_row + 1
        firstcol = 'A'
        # lastcol = chr(ord(firstcol) + len(column_headers))
        
        # notes_coords = firstcol+str(first_notes_row)+':'+lastcol+str(first_notes_row)
        # self.worksheet_facility.merge_range(notes_coords, '', self.formats['notes'])
        self.worksheet_facility.write_rich_string(firstcol+str(first_notes_row)
          , self.formats['superscript'], '1'
          , "The demographic percentages are based on the 2020 Decennial Census' block populations, which are linked to the Census’ 2018-2022 American Community Survey (ACS) five-year demographic averages at the block group level. To derive")
        first_notes_row+=1
        self.worksheet_facility.write_string(firstcol+str(first_notes_row)
          , "  demographic percentages, it is assumed a given block's demographics are the same as the block group in which it is contained. Demographics are tallied for all blocks falling within the indicated radius. ")
        first_notes_row+=1
        self.worksheet_facility.write_rich_string(firstcol+str(first_notes_row)
          , self.formats['superscript'], '2'
          , 'A person is identified as one of six racial/ethnic categories: White, Black, American Indian or Alaska Native, Asian, Other and Multiracial, or Hispanic/Latino. The People of Color population is the total population minus the White population.')
        first_notes_row+=1
        self.worksheet_facility.write_rich_string(firstcol+str(first_notes_row)
          , self.formats['superscript'], '3'
          , 'To avoid double counting, the "Hispanic or Latino" category is treated as a distinct demographic category. A person who identifies as Hispanic or Latino is counted only as Hispanic/Latino for this analysis (regardless of other racial identifiers).')
        first_notes_row+=1
        self.worksheet_facility.write_rich_string(firstcol+str(first_notes_row)        
          , self.formats['superscript'], '4'
          , ('The demographic percentages for people living below the poverty line or below twice the poverty line are based on Census ACS surveys at the block group level that do not include people in group living situations such as'
             ))
        first_notes_row+=1
        self.worksheet_facility.write_string(firstcol+str(first_notes_row)
          , ('dorms, prisons, nursing homes, and military barracks. To derive the nationwide demographic percentages shown, these block group level tallies are summed for all block groups in the nation and then divided by the total U.S. population'
             ))
        first_notes_row+=1
        self.worksheet_facility.write_string(firstcol+str(first_notes_row)
          , ("based on the 2018-2022 ACS. The study area's facility-specific and run group-wide population counts are based on the methodology noted in footnote 1 to derive block-level demographic population counts for the study area,"
             ))
        first_notes_row+=1
        self.worksheet_facility.write_string(firstcol+str(first_notes_row)
          , ('which are then divided by the respective total block-level population (facility-specific and run group-wide) to derive the study area demographic percentages shown.'
             ))
                                                  
        first_notes_row+=1
        self.worksheet_facility.write_rich_string(firstcol+str(first_notes_row)        
          , self.formats['superscript'], '5'
          , ('The demographic percentage for people >= 25 years old without a high school diploma is based on Census ACS data for the total population 25 years old and older at '
             'the block group level, which is used as the denominator when calculating this demographic percentage.'
             ))
        first_notes_row+=1
        self.worksheet_facility.write_rich_string(firstcol+str(first_notes_row)        
          , self.formats['superscript'], '6'
          , ('The Limited English Speaking population is estimated at the block group level by taking the product of the block group population and the fraction of '
             'Limited English Speaking households in the block group, assuming that the number of individuals '))
        first_notes_row+=1
        self.worksheet_facility.write_string(firstcol+str(first_notes_row)
          , ('  per household is the same for Limited English Speaking households '
             'as for the general population, and summed over all block groups.'))
        first_notes_row+=1
        self.worksheet_facility.write_rich_string(firstcol+str(first_notes_row)        
          , self.formats['superscript'], '7'
          , ('The nationwide 2020 Decennial Census population of 334,753,155 is the summation of all Census block populations within the 50 states, the '
             'District of Columbia, and Puerto Rico. Note that the nationwide population based on the'))
        first_notes_row+=1
        self.worksheet_facility.write_string(firstcol+str(first_notes_row)
          , ('  2020 Decennial Census differs slightly from the nationwide population based on the five-year '
             '2018-2022 American Community Survey averages, because the former is not based on a '
             'five-year average.'))
        first_notes_row+=1
        self.worksheet_facility.write_rich_string(firstcol+str(first_notes_row)        
          , self.formats['superscript'], '8'
          , ('The population tally and demographic analysis of the total population surrounding all facilities as a group takes into account neighboring facilities '
             'with overlapping study areas and ensures populations in common are counted only once.'))
        first_notes_row+=1

        # # Set row height for the notes
        # self.worksheet_facility.set_row(first_notes_row-1, 230)


        #------------ Sortable Spreadsheet ----------------------------------------------

        sort_headers = ['Population Basis', 'Longitude', 'Latitude', 'Total Population', 'White', 
                          'People of Color', 'Black',
                          'American Indian or Alaska Native', 'Asian', 'Other and Multiracial', 'Hispanic or Latino',
                          'Age (Years)\n0-17', 'Age (Years)\n18-64', 'Age (Years)\n>=65',
                          'People Living Below the Poverty Level', 
                          'People Living Below Twice the Poverty Level', 
                          'People >= 25 Years Old',
                          'People >= 25 Years Old without a High School Diploma',
                          'People Living in Limited English Speaking Households']
        
        firstcol = 'A'
        lastcol = chr(ord(firstcol) + len(sort_headers))
        top_header_coords = firstcol+'1:'+lastcol+'1'

        # Increase the column width. First column is 16; all others are 12.
        self.worksheet_sort.set_column("A1:A1", 26)
        self.worksheet_sort.set_column("B1:"+lastcol+"1", 12)
              
        # Create column headers
        self.worksheet_sort.set_row(0, 82, self.formats['sub_header_2'])
        for col_num, data in enumerate(sort_headers):
            self.worksheet_sort.write(0, col_num, data)
                
      
        # Add Facility ID, Lat, Lon
        facname_list = self.faclist_df['facility_id'].tolist()
        row_num = 3
        for index, row in self.faclist_df.iterrows():
            self.worksheet_sort.write_string(row_num, 0, row['facility_id'], self.formats['sub_header_6'])
            self.worksheet_sort.write_number(row_num, 1, row['lon'], self.formats['sub_header_3'])
            self.worksheet_sort.write_number(row_num, 2, row['lat'], self.formats['sub_header_3'])
            row_num = row_num + 1
                
        
    def close_workbook(self):
        self.workbook.close()

