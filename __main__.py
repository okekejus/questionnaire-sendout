
# -*- coding: utf-8 -*-
"""
@author: Justin Okeke
Complete LSQ Sendout
"""
import time

# these were the file names on the drive 
import obs_data
import obs_email
import config
import obs_lsq_epds
import config_api


def main():
    print('Importing Access data')
    access_data_dict = obs_data.access_data(path_access = config.path_access)
    obs_email.ObsParticipants.set_access_table(
        access_data_dict['enrolment'], access_data_dict['followup']
    )
    
    # remove appropriate OBS IDs
    ids_exclude = obs_data.access_exclude(
        access_data_dict = access_data_dict,
        excl_previous = False, excl_multiple = False, excl_no_use = True, 
        excl_no_contact = True, excl_no_access = True,
        excl_fetal_demise = True, excl_neonatal_death = True   
    )
    obs_email.ObsParticipants.remove_exclusion_ids(ids_exclude)
    
    obs_email.ObsParticipants.set_emails_link_pwd(
        path_contact = config.path_contact,
        path_link = config.path_link
    )
    
    lsq_dict = {}
    print('Downloading REDCap data')    
    # list of lsq subjects given but not returned
    for lsq_num in range(1, 4):
        lsq_str = 'lsq{}'.format(str(lsq_num))
        # use API to get LSQ info regarding recent completions
        redcap_lsq = obs_data.redcap_lsq_summary(
            config_api.redcap_api[lsq_str], lsq_num
        )
        redcap_lsq = obs_data.lsq_complete(redcap_lsq)
        
        lsq_dict[lsq_str] = obs_email.Lsq(str(lsq_num), redcap_lsq)
        
    print('Updating Access')    
    for key, val in lsq_dict.items():
        val.update_access_returned(
            path_access = config.path_access
        )
      
    print('Determining LSQ statuses')
    for val in lsq_dict.values():
    # get appropriate LSQ statuses
        val.set_lsq_status()
        
    # remove obs ids based on priority of LSQ
    lsq3_ids = [
        obs_id 
        for sublist in lsq_dict['lsq3'].lsq_status.values() 
        for obs_id in sublist
    ]
    lsq_dict['lsq2'].remove_lsq_priority(lsq3_ids)
    lsq_dict['lsq1'].remove_lsq_priority(lsq3_ids)
    
    lsq2_ids = [
        obs_id 
        for sublist in lsq_dict['lsq2'].lsq_status.values() 
        for obs_id in sublist
    ]
    lsq_dict['lsq1'].remove_lsq_priority(lsq2_ids)
    
    print('Sending LSQ emails')
    # num_email tracks how many emails were sent
    num_email = 0
    # send emails, update access
    for val in lsq_dict.values():
        val.send_lsq_emails(10)
        val.update_access_status(path_access = config.path_access)
        for val2 in val.lsq_status.values():
            num_email = num_email + len(val2)
    
    time.sleep(60)    
    # transfer emails from personal to communal email address:
    for i in range(num_email * 2):
        obs_email.transfer_last_email(
            from_email_folder = config.sent_from_email, 
            to_email_folder = config.obs_outlook_folder
        )
    # send confirmation emails
    obs_email.send_email(
        config.notification_email,
        'Successful Weekly Distribution',
        'Number of LSQs sent: ' + str(num_email),
        sent_on_behalf = config.sent_on_behalf
    )
    
    print('Determining EPDS followups')    
    # EPDS info
    lsq2_epds = obs_lsq_epds.Lsq2Epds(
        lsq_dict['lsq2'].update_access_comp['obs_study_id'].tolist(), 
        config_api.redcap_api['lsq2'], cut_off = 10
    ).fu_ids
    lsq3_epds = obs_lsq_epds.Lsq3Epds(
        lsq_dict['lsq3'].update_access_comp['obs_study_id'].tolist(), 
        config_api.redcap_api['lsq3'], cut_off = 10
    ).fu_ids
    lsq_epds = lsq2_epds + lsq3_epds
    
    if len(lsq_epds) > 0:
        epds_body = (
            'The following OBS subjects need to be followed up based on their'
            ' EPDS score:\n\n'
        )
        for id_epds in lsq_epds:
            single_epds = (
                str(id_epds[0]) 
                + ' (EPDS Score: ' + str(id_epds[1]) 
                + '; "Harming myself" answer: ' + id_epds[2]
                + ')\n'
            )
            epds_body = epds_body + single_epds            
    else:
        epds_body = 'Hey Kim, There are no EPDS followups this week\n\n'
    
    obs_email.send_email(
        config.epds_fu_email,
        'EPDS Followup',
        epds_body,
        sent_on_behalf = config.sent_on_behalf
    )

if __name__ == '__main__':
    main()     
