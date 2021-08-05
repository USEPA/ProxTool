import datetime
import os
import re
import numpy as np
import pandas as pd

from com.sca.ca.model.ExcelDataset import ExcelDataset


class FacilityList(ExcelDataset):

    def get_columns(self):
        return ['facility_id', 'lon', 'lat']

    def get_numeric_columns(self):
        return ['lon', 'lat']

    def get_string_columns(self):
        return ['facility_id']