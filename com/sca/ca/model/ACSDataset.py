import datetime
import os
import re
import numpy as np
import pandas as pd

# from com.sca.ca.model.ExcelDataset import ExcelDataset
from com.sca.ca.model.CSVDataset import CSVDataset

class ACSDataset(CSVDataset):

    def get_columns(self):
        return ['bkgrp', 'totalpop', 'p_minority', 'pnh_white', 'pnh_afr_am',
                'pnh_am_ind', 'pnh_othmix', 'pt_hisp', 'p_agelt18', 'p_agegt64',
                'p_2xpov', 'p_pov', 'age_25up', 'p_edulths', 'p_lingiso',
                'age_univ', 'pov_univ', 'edu_univ', 'iso_univ', 'pov_fl', 'edu_fl', 'iso_fl']


    def get_numeric_columns(self):
        return ['totalpop', 'p_minority', 'pnh_white', 'pnh_afr_am',
                'pnh_am_ind', 'pnh_othmix', 'pt_hisp', 'p_agelt18', 'p_agegt64',
                'p_2xpov', 'p_pov', 'age_25up', 'p_edulths', 'p_lingiso',
                'age_univ', 'pov_univ', 'edu_univ', 'iso_univ']


    def get_string_columns(self):
        return ['bkgrp', 'pov_fl', 'edu_fl', 'iso_fl']

