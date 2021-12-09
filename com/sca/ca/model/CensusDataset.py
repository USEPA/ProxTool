import datetime
import os
import re
import numpy as np
from com.sca.ca.model.CSVDataset import CSVDataset


class CensusDataset(CSVDataset):

    def get_columns(self):
        return ['fips', 'blkid', 'population', 'lat', 'lon', 'elev', 'hill', 'urban_pop']

    def get_numeric_columns(self):
        return ['population', 'lat', 'lon', 'elev', 'hill', 'urban_pop']

    def get_string_columns(self):
        return ['fips', 'blkid']