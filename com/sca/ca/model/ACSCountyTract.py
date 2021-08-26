
from com.sca.ca.model.CSVDataset import CSVDataset


class ACSCountyTract(CSVDataset):

    def get_columns(self):
        return ['ID', 'TOTALPOP', 'PCT_MINORITY', 'PCT_WHITE', 'PCT_BLACK', 'PCT_AMIND', 'PCT_OTHER_RACE', 'PCT_HISP',
                'PCT_AGE_LT18', 'PCT_AGE_GT64', 'POV_UNIVERSE', 'PCT_LOWINC', 'PCT_POV', 'EDU_UNIVERSE', 'PCT_EDU_LTHS', 
                'PCT_LINGISO', 'POVERTY_FLAG', 'EDUCATION_FLAG', 'LING_ISO_FLAG']

    def get_numeric_columns(self):
        return ['TOTALPOP', 'PCT_MINORITY', 'PCT_WHITE', 'PCT_BLACK', 'PCT_AMIND', 'PCT_OTHER_RACE', 'PCT_HISP',
                               'PCT_AGE_LT18', 'PCT_AGE_GT64', 'POV_UNIVERSE', 'PCT_LOWINC', 'PCT_POV', 'EDU_UNIVERSE', 'PCT_EDU_LTHS', 
                               'PCT_LINGISO']

    def get_string_columns(self):
        return ['ID', 'POVERTY_FLAG', 'EDUCATION_FLAG', 'LING_ISO_FLAG']



