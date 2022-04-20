
# -*- coding: utf-8 -*-
"""
@author: Justin Okeke
Email sendout
"""


import win32com.client
import pyodbc
import pandas as pd
import re
import numpy as np
import time
from datetime import datetime

import config


def send_email(
    email_address, subject, body, sent_on_behalf = config.sent_on_behalf
):
    """Send email
    
    Parameters
    ----------
    email_address : str
        Email address of the intended recipient
    subject : str
        Subject of email
    body : str
        Body of email
    sent_on_behalf : str, optional
        Sent of behalf of/return email address. The default is 
        config.sent_on_behalf.

    Returns
    -------
    None.

    """                                                               
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = email_address
    mail.SentOnBehalfofName =  sent_on_behalf
    mail.Subject = subject
    mail.body = body
    mail.send

def transfer_last_email(
        from_email_folder = config.sent_from_email, 
        to_email_folder = config.obs_outlook_folder
):
    """Moves last email from original to communal 'Sent Items' folder
    
    Parameters
    ----------
    from_email_folder : str, optional
        Email address associated where emails are sent from. The default is 
        config.sent_from_email.
    to_email_folder : str, optional
        Communal email address where the email moved to ('Sent Items'). The 
        default is config.sent_on_behalf.

    Returns
    -------
    None.

    """
    outlook = (
        win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    )
    personal_sent_folder = (
        outlook.Folders(from_email_folder).Folders('Sent Items')
    )
    communal_sent_folder = (
        outlook.Folders(to_email_folder).Folders('Sent Items')
    )
    personal_sent_folder.Items.GetLast().Move(communal_sent_folder)
    


class ObsParticipants():
    """Associated data for all OBS participants
    
    Attributes
    ----------
    lsq_given_ga: dict
        used to set the gestational age when LSQs are sent; keys are
        associated with an LSQ and the values are associated with the day
        threshold that LSQ gets sent
    lsq_fu_days: int
        number of days after which a follow-up/reminder email will be sent
    status_priorities: dict
        key is LSQ version, value is associated inferior LSQs
    access_table: pandas.dataframe
        Tables from access that have been cleaned and modified.
    id_enrol: list of strings
        List of unique IDs of subjects enrolled in OBS
    access_patient_info: pandas.dataframe
        Contains 'obs_study_id', 'PatientID', 'PatientFirstName',
        'PatientSurname' from Access
    id_excl: list of strings
        List of subjects to exclude
    access_wo_excl: pandas.dataframe
        Access table without the ids to be excluded (id_excl).
    emails_link_pwd_dict:
        key is OBS ID, value is lsq links and associated passwords
    
    """
    # class attributes
    
    lsq_given_ga = config.lsq_given_ga
    lsq_fu_days = config.lsq_fu_days
     
    status_priorities = {
        'LSQ({})Followup3' : [
            'LSQ({})Followup2', 'LSQ({})Followup1', 'LSQ({})Given'
        ],
        'LSQ({})Followup2' : ['LSQ({})Followup1', 'LSQ({})Given'],
        'LSQ({})Followup1' : ['LSQ({})Given'],
    }
    @ classmethod
    def set_access_table(cls, enrolment, followup):
        """Prepare access table and set as class attribute
        
        Parameters
        ----------
        cls : TYPE
            DESCRIPTION.
        enrolment : TYPE
            DESCRIPTION.
        followup : TYPE
            DESCRIPTION.

        Returns
        -------
        None.

        """

        cls.access_table  = cls._clean_access_table(enrolment, followup)
        cls.id_enrol = list(cls.access_table['obs_study_id'].unique())
        
        cls.access_patient_info = cls.access_table.loc[
            :, [
                'obs_study_id', 'PatientID', 
                'PatientFirstName', 'PatientSurname'
            ]
        ].drop_duplicates(
            subset = 'obs_study_id', keep = 'last'
        ).set_index('obs_study_id').to_dict('index')
        

        
    @ classmethod 
    def _clean_access_table(cls, enrolment, followup):
        """Modifies access tables
        
        Combines 'enrolment' and 'followup' tables; changes column types; 
        removes older subjects; adds LMP column
        

        Parameters
        ----------
        enrolment : pandas.dataframe
            Access table containing the enrolment information of subjects
        followup : TYPE
            Access table 

        Returns
        -------
        access_table : pandas.dataframe
            Modified and combined access table

        """
        access_table = pd.merge(
            enrolment.loc[
                :, ['obs_study_id', 'EDC', 'DIPLateEntry', 'DIPCurEnrol']
            ], 
            followup, on = 'obs_study_id'
        )
        
        # change column types
        access_table['EDC'] = pd.to_datetime(
            access_table['EDC'], format = '%Y-%m-%d'
        )
        access_table['obs_study_id'] = access_table['obs_study_id'].astype(str)
        access_table['PatientID'] = (
            access_table['PatientID'].astype(str).str.replace('\.0', '')
        )
        
        # began tracking LSQs May 30, 2016
        access_table = access_table.loc[
            (access_table['EDC'] >= pd.Timestamp(year=2016, month=5, day=30))
        ]

        # add LMP column
        access_table['LMP'] = access_table['EDC'] + pd.DateOffset(days = -280)

        return access_table
        
    @ classmethod   
    def remove_exclusion_ids(cls, id_exclude):
        """Remove OBS IDs from cls.access_table
        
        Parameters
        ----------
        id_exclude : list of str
            IDs to be removed from access table

        Returns
        -------
        None.
        """
        cls.id_excl = id_exclude

        id_enrol_wo_excl = [
            obs_id 
            for obs_id in cls.id_enrol 
            if int(obs_id) not in id_exclude
        ]
        cls.access_wo_excl = cls.access_table[
            cls.access_table['obs_study_id'].isin(id_enrol_wo_excl)
        ]
    
    
    @ classmethod
    def set_emails_link_pwd(cls, path_contact, path_link):
        
        # get email addresses of subjects
        emails = pd.ExcelFile(path_contact).parse('Sheet1') 
        emails['obs_study_id'] = '912' + emails['OBSID'].astype(str)
        emails = emails.loc[:, ['obs_study_id', 'E-mail']]
        # get LSQ links and passwords
        link_pwd = pd.read_csv(path_link)
        link_pwd['obs_study_id'] =  (
            link_pwd['obs_subject_id'].str.replace('-', '', regex = True)
        )
        link_pwd = link_pwd.drop(columns = ['obs_subject_id'])
        # combine email addresses and LSQ links/passwords
        emails_link_pwd = pd.merge(emails, link_pwd, on = 'obs_study_id')
        emails_link_pwd = emails_link_pwd[~emails_link_pwd.isin(cls.id_excl)]
        
        cls.emails_link_pwd_dict = (
            emails_link_pwd.set_index('obs_study_id').to_dict('index')
        )


class Lsq(ObsParticipants):
    """Associated data for all LSQ
    
    
    Attributes
    ----------
    lsq_num: int
        Number associated with particular LSQ; should be 1, 2, or 3
    id_access_not_returned: list
        OBS IDs of subjects who have been given an LSQ but have not returned
        it
    redcap_lsq_comp: pandas.dataframe
        Cleaned REDCap dataframe containing subjects who completed LSQ
        
    update_access_comp: list
        OBS IDs of subjects who have the completion status updated in Access
    lsq_status: dict
        Key contains version of LSQ and value contains list of OBS IDs who
        need to receive associated LSQ

    """


    def __init__(self, lsq_num, redcap_lsq_compl):
        self.lsq_num = str(lsq_num)
        self._set_id_access_not_returned()
        
        self.redcap_lsq_comp = self._clean_redcap_lsq_comp(redcap_lsq_compl) 

    def _set_id_access_not_returned(self):
        """Find given, not returned OBS IDs
        
        Find the OBS IDs of subjects who have been given an LSQ and have not
        completed that LSQ according to Access; set as class attribute
        
        Returns
        -------
        None.

        """
        access_given = self.access_table.loc[
            self.access_table['LSQ({})Given'.format(self.lsq_num)] > 0, 
            'obs_study_id'
        ].tolist()
        access_returned = self.access_table.loc[
            self.access_table['LSQ({})Returned'.format(self.lsq_num)] > 0, 
            'obs_study_id'
        ].tolist()
        self.id_access_not_returned = [
            str(obs_id) 
            for obs_id in access_given 
            if obs_id not in access_returned
        ]
    
    def update_access_returned(self,  path_access):
        """Update Access database with subjects who recently completed LSQ
    
        Parameters
        ----------

        path_access: str
            Path to the Access database
            
        Returns
        -------
        None.                  
        """

        # only get subjects who don't have up to date completion
        self.update_access_comp = self.redcap_lsq_comp[
            self.redcap_lsq_comp['obs_study_id'].isin(
                self.id_access_not_returned
            )
        ]
        if len(self.update_access_comp) > 0:
            self._execute_sql_access(
                    ids_updating = self.update_access_comp,
                    lsq_ver_tmpl = 'LSQ({})Returned', 
                    path_access = path_access
            )

    def _execute_sql_access(self, ids_updating, lsq_ver_tmpl, path_access):
        """Put REDCap data into Access database
        
        Cleans and formats data from REDCap and imports into
        Access database

        Parameters
        ----------
        ids_updating : pandas.dataframe
            OBS IDs and dates that are to be updated in Access
        lsq_ver_tmpl : str
            String associated with a column in the Access database (e.g.
            'LSQ({})Returned')
        path_access : str
            Path to Access database

        Returns
        -------
        None.
        """
        conn =  pyodbc.connect((r'DRIVER={Microsoft Access Driver '
                                    r'(*.mdb, *.accdb)}; DBQ=' + path_access +
                                    r';;UID="";PWD="";'))
        cursor = conn.cursor()
        
        for obs_id, date_compl in dict(ids_updating.values.tolist()).items():
            
            # set data to be put into Access
            patient_id = self.access_patient_info[obs_id]['PatientID']
            patient_first_name = (
                self.access_patient_info[obs_id]['PatientFirstName']
            )
            patient_surname = (
                self.access_patient_info[obs_id]['PatientSurname']
            )
            obs_access_id = re.sub('^912', '912-', obs_id)
            lsq_ver = lsq_ver_tmpl.format(self.lsq_num)
            # put data into Access
            cursor.execute(
                 'INSERT INTO OBSFollowupLog (PatientID, OBSEnrolmentID, '
                 'PatientFirstName, PatientSurname, OBSVisitDate, '
                 '\"{}\") VALUES (?, ?, ?, ?, ?, ?)'.format(lsq_ver),
                 patient_id, obs_access_id, patient_first_name, 
                 patient_surname, date_compl, 1
            )
            conn.commit()
        conn.close()
    
    def _clean_redcap_lsq_comp(self, redcap_lsq_compl):
        """Remove superflous columns from REDCap LSQ completed

        Parameters
        ----------
        redcap_lsq_compl : pandas.dataframe
            REDCap LSQ data of subjects who completed LSQ

        Returns
        -------
            pandas.dataframe with two columns obs_study_id and
            lifestyle_questionnaire_X_timestamp (completion date), where 'X'
            is the LSQ number of interest

        """

        lsq_date_col = 'lifestyle_questionnaire_' + self.lsq_num + '_timestamp'
        redcap_lsq_comp_col = (
            redcap_lsq_compl.loc[:, ['obs_study_id', lsq_date_col]]
        )
        redcap_lsq_comp_col[lsq_date_col] = pd.to_datetime(
            redcap_lsq_comp_col[lsq_date_col], format = '%Y-%m-%d %H:%M:%S'
        ).dt.strftime('%Y-%m-%d').astype(str)
        return redcap_lsq_comp_col
    

    def set_lsq_status(self):
        """Set LSQ status
        
        Returns
        -------
        None.

        """
        # determine OBS IDs of subjects who will not be followed up
        self._set_lsq_no_fu_access()
        self._set_no_fu_days()
        self._lsq_no_fu = list(set(self._lsq_no_fu_access + self._no_fu_days))
        
        # find ALL subjects who should be given an LSQ based on number of days;
        # find ALL subjects who should be followed up based on number of days
        self.lsq_status = {}
        self._set_lsq_status_given()
        self._set_lsq_status_fu()

        self._remove_status_priority()       
        
    def _set_lsq_no_fu_access(self):
        """Find OBS IDs of subjects who don't need followups
        
        Find LSQ IDs who do not need followups based on previous enteries in
        Access database (i.e. returned, followup3, refused, paper or 
        lsq_no_fu_access)
        
        """
        lsq_no_fu_access = self.access_wo_excl.loc[
            (
                (self.access_wo_excl['LSQ(' + self.lsq_num + ')Returned'])
                | (self.access_wo_excl['LSQ(' + self.lsq_num + ')Followup3'])
                | (self.access_wo_excl['LSQ' + self.lsq_num + 'Refused'])
                | (self.access_wo_excl['Paper LSQ' + self.lsq_num])
            ), 'obs_study_id'
        ].unique().tolist()
        
        lsq_no_fu_access.extend(self.redcap_lsq_comp['obs_study_id'].tolist())
        self._lsq_no_fu_access = list(set(lsq_no_fu_access))
        
    def _set_no_fu_days(self):
        """Find subjects where insufficient time has passed since last contact
        
        Returns
        -------
        None.

        """
        no_fu_days_mask = False * len(self.access_wo_excl.index)
        for lsq_status in [
                'LSQ({})Given', 'LSQ({})Followup1', 
                'LSQ({})Followup2', 'LSQ({})Followup3'
        ]:
            temp_mask = (
                (
                    (
                        self.access_wo_excl['OBSVisitDate'] 
                        + pd.DateOffset(days = self.lsq_fu_days)
                    ) > pd.Timestamp.today()
                )
                & (self.access_wo_excl[lsq_status.format(self.lsq_num)])
            )
            no_fu_days_mask = np.logical_or(no_fu_days_mask, temp_mask)

        self._no_fu_days = self.access_wo_excl.loc[
                no_fu_days_mask, 'obs_study_id'
            ].unique().tolist()

    def _set_lsq_status_given(self):
        """Find subjects who have been need to be 'given' the associated LSQ
        
        Returns
        -------
        None.

        """
        if self.lsq_num == '3':
            access_table, given_delivery = self._lsq3_access_delivery()
        else:
            access_table = self.access_wo_excl
            given_delivery = []
        
        # given_ga lists subjects who have passed the LSQ threshold based on 
        # GA in the case of LSQ3, this does not include subjects who have
        # delivered
        given_ga = access_table.loc[
            (
                (
                    access_table['LMP'] 
                    + pd.DateOffset(
                        days = self.lsq_given_ga['lsq' + self.lsq_num + 'ga']
                    )
                ) <= pd.Timestamp.today()
            ), 'obs_study_id'
        ].unique().tolist()
        
        
        given_total = given_delivery + given_ga
        
        # elminate subjects who don't require follow-up
        self.lsq_status['LSQ(' + self.lsq_num + ')Given'] = [
            obs_id
            for obs_id in given_total
            if obs_id not in self._lsq_no_fu
        ]
        
    def _lsq3_access_delivery(self):
        """Find subjects who delivered and passed LSQ3 threshold, and not
        delivered
        

        Returns
        -------
        access_table : pandas.dataframe
            Contains subjects who have who have not delivered
        given_delivery : pandas.dataframe
            Contains subjects who have who have passed the LSQ3 threshold
            based on delivery date
        """
        id_delivery_date = self.access_wo_excl.loc[
            self.access_wo_excl['DeliveryDate'].notna(), 'obs_study_id'
        ].unique().tolist()
        # given_delivery lists subjects who have passed the LSQ3 
        # threshold based on delivery date
        given_delivery = self.access_wo_excl.loc[
            (
                (
                    self.access_wo_excl['DeliveryDate']
                    + pd.DateOffset(
                        days = self.lsq_given_ga['lsq3dd']
                    )
                ) <= pd.Timestamp.today()
            ), 'obs_study_id'
        ].unique().tolist()
        
        # set access_table to those who have not delivered
        access_table = self.access_wo_excl[
            ~self.access_wo_excl['obs_study_id'].isin(id_delivery_date)
        ]
        return access_table, given_delivery
        
    def _set_lsq_status_fu(self):
        """Find subjects who need to be followed up and set as attribute
        
        Returns
        -------
        None.
        """
        status_priorities = [
            'LSQ({})Given', 'LSQ({})Followup1', 
            'LSQ({})Followup2', 'LSQ({})Followup3'
        ]
        for i in range(0, 3):
            elapsed_time_mask = (
                ((
                    self.access_wo_excl['OBSVisitDate'] 
                    + pd.DateOffset(days = self.lsq_fu_days)
                ) <= pd.Timestamp.today())
                & (self.access_wo_excl[
                    status_priorities[i].format(self.lsq_num)
                ])
            )

            fu_total = self.access_wo_excl.loc[
                elapsed_time_mask, 'obs_study_id'
            ].unique().tolist()
            # elminate subjects who don't require follow-up
            self.lsq_status[status_priorities[i + 1].format(self.lsq_num)] = [
                obs_id
                for obs_id in fu_total
                if obs_id not in self._lsq_no_fu
            ]
            
    def _remove_status_priority(self):
        """Remove OBS ID from lower status within LSQ
        
        If OBS ID is in two lists for the same LSQ, method will remove OBS ID
        from the lower priority followup status
        
        Returns
        -------
        None.

        Examples
        --------
        >>> obj.LSQ(1)Given = ['1', '2', '3']
        >>> obj.LSQ(1)Followup1 = ['1', '2']
        >>> obj.LSQ(1)Followup2 = ['1']
        >>> obj._remove_status_priority()
        >>> obj.LSQ(1)Given
        '3'
        >>> obj.LSQ(1)Followup1
        '2'
        >>> obj.LSQ(1)Followup2
        '1'
        """
        for high_pri_status, low_pri_statuses in sorted(
                self.status_priorities.items(), reverse = True
        ):
            high_lsq_ver = str(high_pri_status).format(self.lsq_num)
            for low in low_pri_statuses:
                low_lsq_ver = str(low).format(self.lsq_num)
                self.lsq_status[low_lsq_ver] = [
                    obs_id
                    for obs_id in self.lsq_status[low_lsq_ver]
                    if obs_id not in self.lsq_status[high_lsq_ver]
                ]
    
    def remove_lsq_priority(self, high_pri):
        """Remove OBS ID from all statuses
        
        Takes a list of OBS IDs and removes them from given, followup1,
        followup2, followup3 in associated LSQ; intended to be used between
        LSQs (as opposed to within an LSQ as with _remove_status_priority())
        

        Parameters
        ----------
        high_pri : list of strings
            List of OBS IDs that are higher priority and will be removed

        Returns
        -------
        None.
        
        Examples
        --------
        >>> obj.LSQ(1)Given = ['3']
        >>> obj.LSQ(1)Followup1 = ['2']
        >>> obj.LSQ(1)Followup2 = ['1']
        >>> obj.remove_lsq_priority(['1', '2'])
        >>> obj.LSQ(1)Given = 
        '3'
        >>> obj.LSQ(1)Followup1
        []
        >>> obj.LSQ(1)Followup2
        []

        """
        for low_pri in [
            'LSQ({})Followup3', 'LSQ({})Followup2', 
            'LSQ({})Followup1', 'LSQ({})Given'
        ]:
            low_lsq_ver = str(low_pri).format(self.lsq_num)
            self.lsq_status[low_lsq_ver] = [
                obs_id
                for obs_id in self.lsq_status[low_lsq_ver]
                if obs_id not in high_pri
            ]
    
    def send_lsq_emails(self, delay_sec = 0):
        """Send LSQ link and password emails

        Parameters
        ----------
        delay_sec : int, optional
            Time to delay between emails. The default is 0.

        Returns
        -------
        None.
        """
        for lsq_status, obs_ids in self.lsq_status.items():
            if len(obs_ids) > 0:
                for obs_id in obs_ids:
                    email_address = self.emails_link_pwd_dict[obs_id]['E-mail']
                    # link email
                    lsq_link = (
                        self.emails_link_pwd_dict[obs_id][
                            'lsq{}_website'.format(self.lsq_num)
                        ]
                    )
                    lsq_link_subject = (
                        config.email_info[lsq_status]['link_subject']
                    )
                    lsq_link_body = (
                        config.email_info[lsq_status]['link_body']
                        .format(lsq_link)
                    )
                    send_email(email_address, lsq_link_subject, lsq_link_body)
                    time.sleep(delay_sec)
                    
                    # password email
                    lsq_password = (
                        self.emails_link_pwd_dict[obs_id][
                            'lsq{}_password'.format(self.lsq_num)
                        ]
                    )
                    lsq_link_subject = (
                        config.email_info[lsq_status]['password_subject']
                    )
                    lsq_link_body = (
                        config.email_info[lsq_status]['password_body']
                        .format(lsq_password)
                    )
                    send_email(email_address, lsq_link_subject, lsq_link_body)
                    time.sleep(delay_sec)
   
    def update_access_status(self, path_access):
        """Update Access database with subjects who were sent LSQs
    
        Parameters
        ----------
        redcap_lsq_compl: list of strings
            All subjects how have completed the associated LSQ.
        path_access: str
            Path to the Access database

        """
        for lsq_ver in [
            'LSQ({})Followup3', 'LSQ({})Followup2', 
            'LSQ({})Followup1', 'LSQ({})Given'
        ]:
            lsq_ver_id = self.lsq_status[lsq_ver.format(self.lsq_num)]
            lsq_ver_id = [str(integer) for integer in lsq_ver_id]
            if len(lsq_ver_id) > 0:
                lsq_ver_date = (
                    [str(datetime.today().strftime('%Y-%m-%d'))] * len(lsq_ver_id)
                )
                lsq_ver_for_execution = (
                    pd.DataFrame(list(zip(lsq_ver_id, lsq_ver_date)))
                )
                
                self._execute_sql_access(
                        ids_updating = lsq_ver_for_execution,
                        lsq_ver_tmpl = lsq_ver, 
                        path_access = path_access
                )
