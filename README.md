# Table of Contents
Background
Objective
Tools and Packages
Results


## Background
I was the previous Data Coordinator for the [Ontario Birth Study](http://ontariobirthstudy.com/), a prospective cohort study embedded in clinical practice at Mount Sinai Hospital in Toronto. Birthing parents <17 weeks gestational age (GA) are recruited by the study, and followed into early postpartum (6-10 weeks). 

Study data is collected in the form of 3 lifestyle and 2 diet questionnaires, in addition to abstraction of clinical records from patient charts. Lifestyle questionnaires are hosted on [REDCap](https://www.project-redcap.org/), and are sent to mothers enrolled in the study at 11 weeks GA, 28 weeks GA, and 6 weeks postpartum. 

Manual questionnaire sendout would be impractical, so a programming approach was required. 

## Objective
Design a script that does the following: 
- Filter patients based on exclusion criteria (terminations, withdrawals, no future contact, etc.) 
- Calculate patient GA to determine if survey should be administered 
- Check if survey has been completed, and send reminders at 2 week intervals, for a maximum of 3 reminders 
- Send out surveys and passwords to enrolled patients at appropriate times
- Update SQL database with survey completion/reminder
- Send notification email to OBS Staff indicating completion + summary of completed and sent out surveys.
- Calculate EPDS Score (a tool used to measure postnatal depression) and flag patients with scores >= 13, or a response other than "Never" to questions about self-harm
- Notify OBS Staff if patient has been flagged on EPDS scale 
- Back up database and delete oldest backup. 

## Tools and Packages 

## Results 
Upon my exit from OBS in Feb 2022, the script was still used to complete the above tasks! 
