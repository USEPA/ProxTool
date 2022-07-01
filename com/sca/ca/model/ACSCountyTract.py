
from com.sca.ca.model.CSVDataset import CSVDataset


class ACSCountyTract(CSVDataset):

    def get_columns(self):
        return ['ID', 'totalpop', 'p_minority', 'pnh_white', 'pnh_afr_am',
                'pnh_am_ind', 'pt_hisp', 'pnh_othmix', 'p_agelt18', 'p_agegt64',
                'p_2xpov', 'p_pov', 'p_edulths', 'p_lingiso', 'pov_univ', 'edu_univ', 'iso_univ']

    def get_numeric_columns(self):
        return ['totalpop', 'p_minority', 'pnh_white', 'pnh_afr_am',
                'pnh_am_ind', 'pnh_othmix', 'pt_hisp', 'p_agelt18', 'p_agegt64',
                'pov_univ', 'p_2xpov', 'p_pov', 'edu_univ', 'p_edulths', 'p_lingiso', 'iso_univ']

    def get_string_columns(self):
        return ['ID']



