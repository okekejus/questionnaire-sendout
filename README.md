# Questionnaire Sendout

I am currently the Data Coordinator for the Ontario Birth Study (OBS), a prospective cohort study embedded in practice at Mount Sinai Hospital in Toronto. 


Every week, we have to send our participants Lifestyle Questionnaires (LSQ) which we use to collect self-reported data. These questionnaires (and the associated responses) are housed in the St. Micheal’s Research Electronic Data Capture (REDCap) server.

Questionnaires are sent at 3 time points: early pregnancy (14 weeks gestational), late pregnancy (28 weeks gestational), and early postpartum (4 weeks). If a patient does not complete a questionnaire within two weeks, they are sent a reminder 2 weeks after. This is repeated 2 times after the initial reminder.

The questionnaires patients receive are based on clinically validated instruments, such as the patient health questionnaire (PHQ), generalized anxiety disorder questionnaire (GAD), and Edinburgh postnatal depression scale (EPDS).

OBS ensures the well being of its patients by evaluating EPDS scores weekly. An EPDS score greater than or equal to 13 is a clinical indicator of major depression. The script identifies this too, and I will show how below. The script is also set to send emails updating myself and my coworker on its progress.

This project uses the RODBC and REDCap R packages for the questionnaires, and also makes use of RDCOMClient for emails. 
