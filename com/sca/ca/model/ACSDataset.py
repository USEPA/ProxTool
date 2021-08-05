import datetime
import os
import re
import numpy as np
import pandas as pd

from com.sca.ca.model.ExcelDataset import ExcelDataset


class ACSDataset(ExcelDataset):

    def get_columns(self):
        return ['STCNTRBG', 'TOTALPOP', 'PCT_MINORITY', 'PCT_WHITE', 'PCT_BLACK', 'PCT_AMIND', 'PCT_OTHER_RACE', 'PCT_HISP',
                'PCT_AGE_LT18', 'PCT_AGE_GT64', 'PCT_LOWINC', 'PCT_POV', 'AGE_25UP', 'PCT_EDU_LTHS', 'PCT_LINGISO', 'AGE_UNIVERSE',
                'POV_UNIVERSE', 'EDU_UNIVERSE', 'ISO_UNIVERSE', 'POVERTY_FLAG', 'EDUCATION_FLAG', 'LING_ISO_FLAG']

    def get_numeric_columns(self):
        return ['TOTALPOP', 'PCT_MINORITY', 'PCT_WHITE', 'PCT_BLACK', 'PCT_AMIND', 'PCT_OTHER_RACE', 'PCT_HISP',
                'PCT_AGE_LT18', 'PCT_AGE_GT64', 'PCT_LOWINC', 'PCT_POV', 'AGE_25UP', 'PCT_EDU_LTHS', 'PCT_LINGISO', 'AGE_UNIVERSE',
                'POV_UNIVERSE', 'EDU_UNIVERSE', 'ISO_UNIVERSE']

    def get_string_columns(self):
        return ['STCNTRBG', 'POVERTY_FLAG', 'EDUCATION_FLAG', 'LING_ISO_FLAG']
