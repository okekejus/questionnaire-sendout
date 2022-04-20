# -*- coding: utf-8 -*-
"""
@author: Justin Okeke

Configuring Emails to be sent 

"""

#!/usr/bin/env python
# -*- coding: utf-8 -*-

# email address(es) of OBS staff notified when after the emails are sent
notification_email = 'my_email@something.com'


# email address(es) of OBS staff who will follow-up with subjects who have
# EPDS issues
epds_fu_email = 'email@something.com; email2@something.com'




# email address of individual who is sending the emails - proxy email
sent_from_email = 'my_email@something.com'


# email address of the 'Sent on behalf of'
sent_on_behalf = 'something@something.com'


# OBS communal folder name
obs_outlook_folder = 'Ontario Birth Study (MSH)'


# sets the gestational threshold for when associated LSQ is sent
# e.g. when 'lsq1ga' : (11*7), LSQ1 will be sent after 11 weeks
lsq_given_ga = {
    'lsq1ga': (11*7),
    'lsq2ga': (26*7),
    'lsq3ga': (46*7),
    'lsq3dd': (6*7)
}


# sets the number of days after which a follow-up/reminder email will be sent
lsq_fu_days = (2*7)

# path to the contact list; must be forward slash
path_contact = (
    'T:/filepath/'
    'Patient contact information/Patient contact information.xlsx'
)
# path to file with LSQ links and emails; must be forward slash
path_link = (
    "T:/filepath"
    "Distribution Passwords Emails/LSQpasswords.csv"
)
# path to access database; must be backward slash
path_access = 'T:\filepath\Screening\Research Database.accdb'

# email information to be included in the LSQ emails
email_info = {
    'LSQ(1)Given': {
        'link_subject': ('Ontario Birth Study Lifestyle Questionnaire 1'),
        'link_body':  (
            "Dear Ontario Birth Study participant:\n\nThank you for enrolling"
            " in the Ontario Birth Study. As part of the study, you will be"
            " asked to complete 3 Lifestyle Questionnaires and 1 Diet History"
            " Questionnaire. These questionnaires are designed to improve"
            " our understanding of mother's and infant's health during"
            " pregnancy and how this influences health over the life"
            " course.\nA link to Lifestyle Questionnaire 1 is below:\n\n"
            "{}"  # link place holder
            " \n\nA separate email will be sent to this email address which"
            " contains a password to access your personal Lifestyle"
            " Questionnaire 1. \n\nYou will receive Lifestyle Questionnaire"
            " 2 when you are approximately 28 weeks gestational age. If you"
            " have any questions, feel free to contact the Ontario Birth"
            " Study at ontariobirthstudy@mtsinai.on.ca or 416-586-4800 ext."
            " 6036. Once again, thank you for participating in the Ontario"
            " Birth Study. Without your help, this research would not be"
            " possible.\n\nSincerely,\n\nThe Ontario Birth Study team"
        ),
        'password_subject': (
            'Ontario Birth Study Lifestyle Questionnaire 1 Password'
        ),
        'password_body': (
            "Dear Ontario Birth Study participant:\n\nYour password for"
            " Lifestyle Questionnaire 1 is as follows:\n\n"
            "{}"
            " \n\nPlease click the link provided in the previous email and"
            " enter the above password to access your personal Lifestyle"
            " Questionnaire.\n\nSincerely,\n\nThe Ontario Birth Study team"
        )
    },

    'LSQ(1)Followup1': {
        'link_subject': (
            "Ontario Birth Study Lifestyle Questionnaire 1 Follow-up"
        ),
        'link_body': (
            "Dear Ontario Birth Study participant:\n\nThis is just a reminder"
            " that you have not completed your most recent Lifestyle"
            " Questionnaire.\n\nA link to Lifestyle Questionnaire 1 is"
            " below:\n\n"
            "{}"
            " \n\nA separate email will be sent to this email address which"
            " contains a password to access your personal Lifestyle"
            " Questionnaire 1. \n\nIf you have any questions, feel free to"
            " contact the Ontario Birth Study at"
            " ontariobirthstudy@mtsinai.on.ca or 416-586-4800 ext. 6036."
            " \n\nSincerely,\n\nThe Ontario Birth Study team"
        ),
        'password_subject': (
            'Ontario Birth Study Lifestyle Questionnaire 1 Password'
        ),
        'password_body': (
            "Dear Ontario Birth Study participant:\n\nYour password for"
            " Lifestyle Questionnaire 1 is as follows:\n\n"
            "{}"
            " \n\nPlease click the link provided in the previous email and"
            " enter the above password to access your personal Lifestyle"
            " Questionnaire.\n\nSincerely,\n\nThe Ontario Birth Study team"
        )
    },

    'LSQ(1)Followup2': {
        'link_subject': (
            'Ontario Birth Study Lifestyle Questionnaire 1 Follow-up'
        ),
        'link_body': (
            "Dear Ontario Birth Study participant:\n\nThis is just a reminder"
            " that you have not completed your most recent Lifestyle"
            " Questionnaire.\n\nA link to Lifestyle Questionnaire 1 is"
            " below:\n\n"
            "{}"
            " \n\nA separate email will be sent to this email address which"
            " contains a password to access your personal Lifestyle"
            " Questionnaire 1. \n\nIf you would like a paper copy of the"
            " Lifestyle Questionnaire or if you have any questions, feel free"
            " to contact the Ontario Birth Study at"
            " ontariobirthstudy@mtsinai.on.ca or 416-586-4800 ext. 6036."
            " \n\nSincerely,\n\nThe Ontario Birth Study team"
        ),
        'password_subject': (
            "Ontario Birth Study Lifestyle Questionnaire 1 Password"
        ),
        'password_body': (
            "Dear Ontario Birth Study participant:\n\nYour password for"
            " Lifestyle Questionnaire 1 is as follows:\n\n"
            "{}"
            " \n\nPlease click the link provided in the previous email and"
            " enter the above password to access your personal Lifestyle"
            " Questionnaire.\n\nSincerely,\n\nThe Ontario Birth Study team"
        )
    },

    'LSQ(1)Followup3': {
        'link_subject': (
            "Ontario Birth Study Lifestyle Questionnaire 1 Follow-up"
        ),
        'link_body': (
            "Dear Ontario Birth Study participant:\n\nThis is just a reminder"
            " that you have not completed your most recent Lifestyle"
            " Questionnaire.\n\nA link to Lifestyle Questionnaire 1 is"
            " below:\n\n"
            "{}"
            " \n\nA separate email will be sent to this email address which"
            " contains a password to access your personal Lifestyle"
            " Questionnaire 1. \n\nIf you would like a paper copy of the"
            " Lifestyle Questionnaire or if you have any questions,"
            " feel free to contact the Ontario Birth Study at"
            " ontariobirthstudy@mtsinai.on.ca or 416-586-4800 ext. 6036."
            " \n\nSincerely,\n\nThe Ontario Birth Study team"
        ),
        'password_subject': (
            "Ontario Birth Study Lifestyle Questionnaire 1 Password"
        ),
        'password_body': (
            "Dear Ontario Birth Study participant:\n\nYour password for"
            " Lifestyle Questionnaire 1 is as follows:\n\n"
            "{}"
            " \n\nPlease click the link provided in the previous email and"
            " enter the above password to access your personal Lifestyle"
            " Questionnaire.\n\nSincerely,\n\nThe Ontario Birth Study team"
        )
    },

    'LSQ(2)Given': {
        'link_subject': ("Ontario Birth Study Lifestyle Questionnaire 2"),
        'link_body': (
            "Dear Ontario Birth Study participant:\n\nIncluded below is a"
            " link to the Ontario Birth Study Lifestyle Questionnaire 2:\n\n"
            "{}"
            " \n\nA separate email will be sent to this email address which"
            " contains a password to access your personal Lifestyle"
            " Questionnaire 2. \n\nYou will receive Lifestyle Questionnaire 3"
            " when you are approximately 6 weeks postpartum. If you have any"
            " questions, feel free to contact the Ontario Birth Study at"
            " ontariobirthstudy@mtsinai.on.ca or 416-586-4800 ext. 6036."
            " \n\nOnce again, thank you for participating in the Ontario"
            " Birth Study. Without your help, this research would not be"
            " possible.\n\nSincerely,\n\nThe Ontario Birth Study team"
        ),
        'password_subject': (
            "Ontario Birth Study Lifestyle Questionnaire 2 Password"
        ),
        'password_body': (
            "Dear Ontario Birth Study participant:\n\nYour password for"
            " Lifestyle Questionnaire 2 is as follows:\n\n"
            "{}"
            " \n\nPlease click the link provided in the previous email and"
            " enter the above password to access your personal Lifestyle"
            " Questionnaire.\n\nSincerely,\n\nThe Ontario Birth Study team"
        )
    },

    'LSQ(2)Followup1': {
        'link_subject': (
            'Ontario Birth Study Lifestyle Questionnaire 2 Follow-up'
        ),
        'link_body': (
            "Dear Ontario Birth Study participant:\n\nThis is just a reminder"
            " that you have not completed your most recent Lifestyle"
            " Questionnaire.\n\nA link to Lifestyle Questionnaire 2 is"
            " below:\n\n"
            "{}"
            " \n\nA separate email will be sent to this email address which"
            " contains a password to access your personal Lifestyle"
            " Questionnaire 2. \n\nIf you have any questions, feel free to"
            " contact the Ontario Birth Study at"
            " ontariobirthstudy@mtsinai.on.ca or 416-586-4800 ext. 6036."
            " \n\nSincerely,\n\nThe Ontario Birth Study team"
        ),
        'password_subject': (
            "Ontario Birth Study Lifestyle Questionnaire 2 Password"
        ),
        'password_body': (
            "Dear Ontario Birth Study participant:\n\nYour password for"
            " Lifestyle Questionnaire 2 is as follows:\n\n"
            "{}"
            " \n\nPlease click the link provided in the previous email and"
            " enter the above password to access your personal Lifestyle"
            " Questionnaire.\n\nSincerely,\n\nThe Ontario Birth Study team"
        )
    },

    'LSQ(2)Followup2': {
        'link_subject': (
            "Ontario Birth Study Lifestyle Questionnaire 2 Follow-up"
        ),
        'link_body': (
            "Dear Ontario Birth Study participant:\n\nThis is just a reminder"
            " that you have not completed your most recent Lifestyle"
            " Questionnaire.\n\nA link to Lifestyle Questionnaire 2 is"
            " below:\n\n"
            "{}"
            " \n\nA separate email will be sent to this email address which"
            " contains a password to access your personal Lifestyle"
            " Questionnaire 2. \n\nIf you would like a paper copy of the"
            " Lifestyle Questionnaire or if you have any questions, feel free"
            " to contact the Ontario Birth Study at"
            " ontariobirthstudy@mtsinai.on.ca or 416-586-4800 ext. 6036."
            " \n\nSincerely,\n\nThe Ontario Birth Study team"
        ),
        'password_subject': (
            "Ontario Birth Study Lifestyle Questionnaire 2 Password"
        ),
        'password_body': (
            "Dear Ontario Birth Study participant:\n\nYour password for"
            " Lifestyle Questionnaire 2 is as follows:\n\n"
            "{}"
            " \n\nPlease click the link provided in the previous email and"
            " enter the above password to access your personal Lifestyle"
            " Questionnaire.\n\nSincerely,\n\nThe Ontario Birth Study team"
        )
    },

    'LSQ(2)Followup3': {
        'link_subject': (
            "Ontario Birth Study Lifestyle Questionnaire 2 Follow-up"
        ),
        'link_body': (
            "Dear Ontario Birth Study participant:\n\nThis is just a reminder"
            " that you have not completed your most recent Lifestyle"
            " Questionnaire.\n\nA link to Lifestyle Questionnaire 2"
            " is below:\n\n"
            "{}"
            " \n\nA separate email will be sent to this email address which"
            " contains a password to access your personal Lifestyle"
            " Questionnaire 2. \n\nIf you would like a paper copy of the"
            " Lifestyle Questionnaire or if you have any questions, feel"
            " free to contact the Ontario Birth Study at"
            " ontariobirthstudy@mtsinai.on.ca or 416-586-4800 ext. 6036."
            " \n\nSincerely,\n\nThe Ontario Birth Study team"
        ),
        'password_subject': (
            "Ontario Birth Study Lifestyle Questionnaire 2 Password"
        ),
        'password_body': (
            "Dear Ontario Birth Study participant:\n\nYour password for"
            " Lifestyle Questionnaire 2 is as follows:\n\n"
            "{}"
            " \n\nPlease click the link provided in the previous email and"
            " enter the above password to access your personal Lifestyle"
            " Questionnaire.\n\nSincerely,\n\nThe Ontario Birth Study team"
        )
    },

    'LSQ(3)Given': {
        'link_subject': ("Ontario Birth Study Lifestyle Questionnaire 3"),
        'link_body': (
            "Dear Ontario Birth Study participant: \n\nCongratulations on"
            " your new baby! This is the last survey for the Ontario Birth"
            " Study. Once you complete this survey, you have fulfilled your"
            " obligations to this study. A link to the Lifestyle"
            " Questionnaire 3 is below:\n\n"
            "{}"
            " \n\nA separate email will be sent to this email address which"
            " contains a password to access your personal Lifestyle"
            " Questionnaire 3. \n\nIf you have any questions, feel free to"
            " contact the Ontario Birth Study at"
            " ontariobirthstudy@mtsinai.on.ca or 416-586-4800 ext. 6036."
            " \n\nOnce again, thank you for participating in the Ontario"
            " Birth Study. Without your help, this research would not be"
            " possible.\n\nSincerely,\n\nThe Ontario Birth Study team"
        ),
        'password_subject': (
            "Ontario Birth Study Lifestyle Questionnaire 3 Password"
        ),
        'password_body': (
            "Dear Ontario Birth Study participant:\n\nYour password for"
            " Lifestyle Questionnaire 3 is as follows:\n\n"
            "{}"
            " \n\nPlease click the link provided in the previous email and"
            " enter the above password to access your personal Lifestyle"
            " Questionnaire.\n\nSincerely,\n\nThe Ontario Birth Study team"
        )
    },

    'LSQ(3)Followup1': {
        'link_subject': (
            "Ontario Birth Study Lifestyle Questionnaire 3 Follow-up"
        ),
        'link_body': (
            "Dear Ontario Birth Study participant:\n\nThis is just a reminder"
            " that you have not completed your most recent Lifestyle"
            " Questionnaire.\n\nA link to Lifestyle Questionnaire 3 is"
            " below:\n\n"
            "{}"
            " \n\nA separate email will be sent to this email address which"
            " contains a password to access your personal Lifestyle"
            " Questionnaire 3. \n\nIf you have any questions, feel free to"
            " contact the Ontario Birth Study at"
            " ontariobirthstudy@mtsinai.on.ca or 416-586-4800 ext. 6036."
            " \n\nSincerely,\n\nThe Ontario Birth Study team"
        ),
        'password_subject': (
            "Ontario Birth Study Lifestyle Questionnaire 3 Password"
        ),
        'password_body': (
            "Dear Ontario Birth Study participant:\n\nYour password for"
            " Lifestyle Questionnaire 3 is as follows:\n\n"
            "{}"
            " \n\nPlease click the link provided in the previous email and"
            " enter the above password to access your personal Lifestyle"
            " Questionnaire.\n\nSincerely,\n\nThe Ontario Birth Study team"
        )
    },

    'LSQ(3)Followup2': {
        'link_subject': (
            "Ontario Birth Study Lifestyle Questionnaire 3 Follow-up"
        ),
        'link_body': (
            "Dear Ontario Birth Study participant:\n\nThis is just a reminder"
            " that you have not completed your most recent Lifestyle"
            " Questionnaire.\n\nA link to Lifestyle Questionnaire 3 is"
            " below:\n\n"
            "{}"
            " \n\nA separate email will be sent to this email address which"
            " contains a password to access your personal Lifestyle"
            " Questionnaire 3. \n\nIf you would like a paper copy of the"
            " Lifestyle Questionnaire or if you have any questions, feel free"
            " to contact the Ontario Birth Study at"
            " ontariobirthstudy@mtsinai.on.ca or 416-586-4800 ext. 6036."
            " \n\nSincerely,\n\nThe Ontario Birth Study team"
        ),
        'password_subject': (
            "Ontario Birth Study Lifestyle Questionnaire 3 Password"
        ),
        'password_body': (
            "Dear Ontario Birth Study participant:\n\nYour password for"
            " Lifestyle Questionnaire 3 is as follows:\n\n"
            "{}"
            " \n\nPlease click the link provided in the previous email and"
            " enter the above password to access your personal Lifestyle"
            " Questionnaire.\n\nSincerely,\n\nThe Ontario Birth Study team"
        )
    },

    'LSQ(3)Followup3': {
        'link_subject': (
            "Ontario Birth Study Lifestyle Questionnaire 3 Follow-up"
        ),
        'link_body': (
            "Dear Ontario Birth Study participant:\n\nThis is just a reminder"
            " that you have not completed your most recent Lifestyle"
            " Questionnaire.\n\nA link to Lifestyle Questionnaire 3 is"
            " below:\n\n"
            "{}"
            " \n\nA separate email will be sent to this email address which"
            " contains a password to access your personal Lifestyle"
            " Questionnaire 3. \n\nIf you would like a paper copy of the"
            " Lifestyle Questionnaire or if you have any questions, feel free"
            " to contact the Ontario Birth Study at"
            " ontariobirthstudy@mtsinai.on.ca or 416-586-4800 ext. 6036."
            " \n\nSincerely,\n\nThe Ontario Birth Study team"
        ),
        'password_subject': (
            "Ontario Birth Study Lifestyle Questionnaire 3 Password"
        ),
        'password_body': (
            "Dear Ontario Birth Study participant:\n\nYour password for"
            " Lifestyle Questionnaire 3 is as follows:\n\n"
            "{}"
            " \n\nPlease click the link provided in the previous email and"
            " enter the above password to access your personal Lifestyle"
            " Questionnaire.\n\nSincerely,\n\nThe Ontario Birth Study team"
        )
    },
}
