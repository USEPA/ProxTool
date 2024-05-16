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
                'pnh_am_ind', 'pnh_asian', 'pt_hisp', 'pnh_othmix', 'p_agelt18', 'p_agegt64',
                'p_2xpov', 'p_pov', 'p_edulths', 'p_lingiso',
                'pov_univ', 'edu_univ', 'iso_univ']


    def get_numeric_columns(self):
        return ['totalpop', 'p_minority', 'pnh_white', 'pnh_afr_am',
                'pnh_am_ind', 'pnh_asian', 'pnh_othmix', 'pt_hisp', 'p_agelt18', 'p_agegt64',
                'p_2xpov', 'p_pov', 'p_edulths', 'p_lingiso',
                'pov_univ', 'edu_univ', 'iso_univ']


    def get_string_columns(self):
        return ['bkgrp']

