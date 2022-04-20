
# -*- coding: utf-8 -*-
"""
@author: Justin Okeke
Retrieve relevant data for questionnaire sendout
"""


import pyodbc
import pandas as pd
import requests
import io
import re


def access_data(
    path_access, enrolment = True, followup = True, screening = False
):
    """Get data from Access database
    

    Parameters
    ----------
    path_access : str, optional
        Path to Access database.
    enrolment : bool, optional
        Setting to True will return dictionary containing 'enrolment' 
        (i.e. 'OBS Enrolment Log') pandas dataframe. The default is True.
    followup : bool, optional
        Setting to True will return dictionary containing 'enrolment' 
        (i.e. 'OBSFollowupLog') pandas dataframe. The default is True.
    screening : bool, optional
        Setting to True will return dictionary containing 'enrolment' 
        (i.e. 'OBS Screening Log') pandas dataframe. The default is False.

    Returns
    -------
    dict of pandas.dataframes
        'enrolment' contains 'OBS Enrolment Log' Access table
        'followup' contains 'OBSFollowupLog' Access table
        'screening' contains 'OBS Screening Log' Access table    
    """
   
    cnxn = pyodbc.connect((r'DRIVER={Microsoft Access Driver '
                            r'(*.mdb, *.accdb)}; DBQ=' + path_access +
                            r';;UID="";PWD="";'))
    
    access_dict = {}
    if enrolment:
        access_dict['enrolment'] = pd.read_sql(
            "Select * From [OBS Enrolment Log]", cnxn
        )
    if followup:
        access_dict['followup'] = pd.read_sql(
            "Select * From [OBSFollowupLog]", cnxn
        )
    if screening:
        access_dict['screening'] = pd.read_sql(
            "Select * From [OBS Screening log]", cnxn
        )
    cnxn.close()
    
    
    for key, value in access_dict.items():
        access_dict[key] = clean_access_table(value)

    return access_dict

   
def clean_access_table(access_table):
    """Clean Access Table
    

    Parameters
    ----------
    access_table : pandas.DataFrame
        DESCRIPTION.

    Returns
    -------
    access_table : pandas.DataFrame
        DESCRIPTION.

    """
    access_table = access_table.drop(
        access_table[access_table['OBSEnrolmentID'].isna()
    ].index)
    
    access_table['OBSEnrolmentID'] = (
        access_table['OBSEnrolmentID'].replace('[^0-9]', '', regex = True)
        .astype(int)
    )
    #access_table['OBSEnrolmentID'] = access_table['OBSEnrolmentID'].astype(int)
    access_table = access_table.rename(
        columns = {'OBSEnrolmentID':'obs_study_id'}
    )
        
    return access_table


# remove defaults for previous and multiples in access_exclude
def access_exclude(
    access_data_dict, excl_previous = True, excl_multiple = True, 
    excl_no_use = True, excl_no_contact = False, excl_no_access = False, 
    excl_fetal_demise = False, excl_neonatal_death = False       
):
    """ Get list of OBS subject IDs to exclude
    
    Parameters
    ----------
    access_data_dict : dict of pandas dataframes
        Expected to be derived from access_data(enrolment = True, 
        followup = True).
    excl_previous : bool, optional
        Setting to True will return most recent OBS subject id(s) for 
        previously enrolled OBS subjects. The default is True.
    excl_multiple : bool, optional
        Setting to True will return OBS subject ids of subjects who had 
        multiples. The default is True.
    excl_no_use : bool, optional
        Setting to True will return OBS subject ids who choose a "No Further 
        Use" withdrawal. The default is True.
    excl_no_contact : bool, optional
        Setting to True will return OBS subject ids who choose a "No Further 
        Contact" withdrawal. The default is True.
    excl_no_access : bool, optional
        Setting to True will return OBS subject ids who choose a "No Further 
        Access" withdrawal. The default is False.
    excl_fetal_demise : bool, optional
        Setting to True will return OBS subject ids who had a fetal demise 
        during the associated OBS enrolment. The default is False.
    excl_neonatal_death : bool, optional
        Setting to True will return OBS subject ids who had a neonatal death
        during the associated OBS enrolment. The default is False.
    

    Returns
    -------
    access_exclude : list
        OBS subjects to be excluded

    """

    # needs to include most up to date acceess database
  
#    if access_data_dict == None:
#        access_data_dict = access_data()
    followup = access_data_dict['followup']
    enrolment = access_data_dict['enrolment']

    excl_type_dict = {'NoUse': excl_no_use,
                      'NoContact': excl_no_contact,
                      'NoAccess': excl_no_access,
                      'Fetal Demise/Termination': excl_fetal_demise,
                      'Neonatal death': excl_neonatal_death}

    access_exclude = []
    for key, value in excl_type_dict.items():
        if value:
            access_exclude.extend(
                followup.loc[followup[key] == 1, 'obs_study_id'].tolist()
            )

    if excl_previous:
        access_exclude = access_exclude.extend(
            enrolment.loc[enrolment['Previous OBS participant'], 
                'obs_study_id'
            ].tolist()
        )
        
    if excl_multiple:
        access_exclude = access_exclude.extend(
            followup.loc[followup['TwinBDelivery'].notnull(), 
                'obs_study_id'
            ].tolist()
        )
    
    access_exclude = list(set(access_exclude))
    
    return access_exclude



#REDCap data

# for clinic, need to get a different way
def redcap_data(api_param, add_param = {}):
    """Get data from AHRC REDCap
    

    Parameters
    ----------
    api_param : str
        API token for associated REDCap data.
    add_param : TYPE, optional
        Add or overwrite request parameters. The default is {}.

    Returns
    -------
    rc_data : pandas.DataFrame
        REDCap data.

    """
    api_req_data = {
        'content': 'record',
        'format': 'csv',
        'type': 'flat',
        'rawOrLabel': 'raw',
        'rawOrLabelHeaders': 'raw',
        'exportCheckboxLabel': 'false',
        'exportSurveyFields': 'true', #necessary for completion date
        'returnFormat': 'json',
        'exportDataAccessGroups': 'true',
    }
    api_req_data.update({'token': api_param})
    api_req_data.update(add_param)
    
    req = requests.post(
        'https://redcap.smh.ca/redcap/api/', api_req_data
    ).content

    rc_data = pd.read_csv(
        io.StringIO(req.decode('utf-8')),  
        sep = ',',
        error_bad_lines = False,
        index_col = False,
        dtype = str
    )
    return rc_data

def redcap_lsq_summary(api_param, version):
    """Get REDCap LSQ summary data
    

    Parameters
    ----------
    api_param : str
        DESCRIPTION.
    version : int or str
        Version of LSQ requesting (e.g. '1' for LSQ1).

    Returns
    -------
    lsq_summary : pandas.DataFrame
        REDCap LSQ summary data.

    """
    
    lsq_summary = redcap_data(
        api_param,
        {
            'fields[0]' : 'obs_study_id',
             'fields[1]' : (
                 'lifestyle_questionnaire_' 
                 + re.sub(r'^[a-zA-Z]*', '', str(version)) 
                 + '_timestamp'
            ),
            'fields[2]' : (
                'lifestyle_questionnaire_' 
                + re.sub(r'^[a-zA-Z]*', '', str(version))
                + '_complete'),
        }
    )
    
    return lsq_summary
    
    

def lsq_complete(lsq):
    """Clean LSQ dataframe
    

    Parameters
    ----------
    lsq : pandas.DataFrame
        LSQ data derived from redcap_data()

    Returns
    -------
    complete_lsq : pandas.DataFrame
        DataFrame only containing completed LSQs with 'password' and
        'redcap_survey_identifier' columns removed

    """
        
    for lsq_version in range(1, 4):
        complete_col = (
            'lifestyle_questionnaire_' + str(lsq_version) +'_complete'
        )
        if complete_col in list(lsq.columns.values):
            complete_lsq = lsq.loc[lsq[complete_col] == '2']
            complete_lsq = complete_lsq.drop(columns = [
                'password', 'redcap_survey_identifier', 
            ], errors = 'ignore')
            break
    
    return complete_lsq
    

def redcap_clinic(api_param):
    """Get AHRC REDCap clinic data
    
    Parameters
    ----------
    api_param : str
        API token for associated REDCap data.
        
    Returns
    -------
    complete_lsq : pandas.DataFrame
        DataFrame only containing REDCap clinic data.
        
    Notes
    -----
        Clinical data requires special treament due to export size 
        restrictions. Creates post request for each obs id and merges.
    
    """

    rc_data = redcap_data(api_param, {'fields[0]': 'obs_id'})
    
    obs_id_lst = list(rc_data['obs_id'].unique())
    
    clinic_df_lst = []
    for obs_id in obs_id_lst:
        rc_data = redcap_data(api_param, {'records[0]': obs_id})
        clinic_df_lst.append(rc_data)    
    merge_rc = pd.concat(clinic_df_lst)
    
    return merge_rc 

