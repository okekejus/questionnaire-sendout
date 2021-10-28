# LSQ SENDOUT

## Prepping libraries to be used ===================================================================
library('pacman')

pacman::p_load(pacman, RODBC, RDCOMClient, openxlsx)

# Pacman: package manager, allows you to load multiple packages 
# RODBC: allows you to reach into Access database for patient information 
# RDCOMClient: sends emails to participants, me, Kim 


## Modified redcap_read_oneshot from REDCapR library ===============================================
redcap_read_oneshot <- function (redcap_uri, 
                                 token, 
                                 records = NULL, 
                                 records_collapsed = "",
                                 fields = NULL, 
                                 fields_collapsed = "", 
                                 events = NULL, 
                                 events_collapsed = "",
                                 export_data_access_groups = FALSE, 
                                 raw_or_label = "raw",
                                 verbose = TRUE, 
                                 config_options = NULL)
{
  start_time <- Sys.time()
  if (missing(redcap_uri))
    stop("The required parameter `redcap_uri` was missing from the call to `redcap_read_oneshot()`.")
  if (missing(token))
    stop("The required parameter `token` was missing from the call to `redcap_read_oneshot()`.")
  if (all(nchar(records_collapsed) == 0))
    records_collapsed <- ifelse(is.null(records), "", paste0(records,
                                                             collapse = ","))
  if ((length(fields_collapsed) == 0L) | is.null(fields_collapsed) |
      all(nchar(fields_collapsed) == 0L))
    fields_collapsed <- ifelse(is.null(fields), "", paste0(fields,
                                                           collapse = ","))
  if (all(nchar(events_collapsed) == 0))
    events_collapsed <- ifelse(is.null(events), "", paste0(events,
                                                           collapse = ","))
  export_data_access_groups_string <- ifelse(export_data_access_groups,
                                             "true", "false")
  post_body <- list(token = token, content = "record", format = "csv",
                    type = "flat", rawOrLabel = raw_or_label, exportDataAccessGroups = export_data_access_groups_string,
                    records = records_collapsed, fields = fields_collapsed,
                    events = events_collapsed, exportSurveyFields = "true")
  result <- httr::POST(url = redcap_uri, body = post_body,
                       config = config_options)
  status_code <- result$status
  LSQdownload <<- status_code
  success <- (status_code == 200L)
  raw_text <- httr::content(result, "text")
  elapsed_seconds <- as.numeric(difftime(Sys.time(), start_time,
                                         units = "secs"))
  regex_cannot_connect <- "^The hostname \\((.+)\\) / username \\((.+)\\) / password \\((.+)\\) combination could not connect.+"
  if (any(grepl(regex_cannot_connect, raw_text)))
    success <- FALSE
  if (success) {
    try({
      ds <- utils::read.csv(text = raw_text, stringsAsFactors = FALSE)
    }, silent = TRUE)
    if (exists("ds") & (class(ds) == "data.frame")) {
      outcome_message <- paste0(format(nrow(ds), big.mark = ",",
                                       scientific = FALSE, trim = TRUE), " records and ",
                                format(length(ds), big.mark = ",", scientific = FALSE,
                                       trim = TRUE), " columns were read from REDCap in ",
                                round(elapsed_seconds, 1), " seconds.  The http status code was ",
                                status_code, ".")
      raw_text <- ""
    }
    else {
      success <- FALSE
      ds <- data.frame()
      outcome_message <- paste0("The REDCap read failed.  The http status code was ",
                                status_code, ".  The 'raw_text' returned was '",
                                raw_text, "'.")
    }
  }
  else {
    ds <- data.frame()
    outcome_message <- paste0("The REDCapR read/export operation was not successful.  The error message was:\n",
                              raw_text)
  }
  if (verbose)
    message(outcome_message)
  return(list(data = ds, success = success, status_code = status_code,
              outcome_message = outcome_message, records_collapsed = records_collapsed,
              fields_collapsed = fields_collapsed, events_collapsed = events_collapsed,
              elapsed_seconds = elapsed_seconds, raw_text = raw_text))
}


## First Reach into accdb, to check which LSQs have been completed, and by who =====================
myconnection <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};
                                  DBQ= filepath")
enrolment <- sqlFetch(myconnection, "OBS Enrolment Log")
followup <- sqlFetch(myconnection, "OBS Followup Log")
close(myconnection)

enrolment <- enrolment[, c("OBSEnrolmentID", "EDC")]
total <- merge(enrolment, followup, by="OBSEnrolmentID")
total$OBSVisitDate <- as.Date(total$OBSVisitDate)
total$DeliveryDate <- as.Date(total$DeliveryDate)
total$EDC <- as.Date(total$EDC)

#drop all subjects who have delivered more than 14 weeks ago
total <- total[total$EDC >= as.Date("05/30/16", "%m/%d/%y"), ]
total <- total[!is.na(total$OBSEnrolmentID), ]
OBSID <- as.character(unique(total$OBSEnrolmentID))

#function to see if patient gave consent or didn't withdrawl
invisible(sapply(OBSID, function(OBSEnrolmentID){
  enrol_table <- total[total$OBSEnrolmentID == OBSEnrolmentID, ]
  if(sum(enrol_table$OBSPatientWithdrawl) > 0 |
     sum(enrol_table[, "Fetal Demise/Termination"]) > 0 |
     sum(enrol_table[, "Neonatal death"]) > 0 ){ #add more variables which may eliminate subjects
    OBSID <<- OBSID[!OBSID %in% c(OBSEnrolmentID)]
  }
}))


## Creates check_if_completed data frame + fills it  with OBS ids and associated LSQs not completed=========

check_if_completed_1 <- c()
check_if_completed_2 <- c()
check_if_completed_3 <- c()

invisible(sapply(OBSID, function(enrolmentID){
 
  enrol_table <- total[total$OBSEnrolmentID == enrolmentID, ]
  if(sum(enrol_table[, "LSQ(1)Given"]) > 0 &
     sum(enrol_table[, "LSQ(1)Returned"]) == 0) {
    check_if_completed_1 <<- c(check_if_completed_1, enrolmentID)
  }
  if(sum(enrol_table[, "LSQ(2)Given"]) > 0 &
     sum(enrol_table[, "LSQ(2)Returned"]) == 0) {
    check_if_completed_2 <<- c(check_if_completed_2, enrolmentID)
  }
  if(sum(enrol_table[, "LSQ(3)Given"]) > 0 &
     sum(enrol_table[, "LSQ(3)Returned"]) == 0) {
    check_if_completed_3 <<- c(check_if_completed_3, enrolmentID)
  }
}))

repeat{
  
## Use the API to check if surveys have been completed (REDCapR) =============================================
# LSQ1 
secret_token <- "token"
api_url <- "api link"
LSQ1complete <- redcap_read_oneshot(redcap_uri = api_url,
                                      token = secret_token,
                                      fields =  c("obs_study_id", 
                                                  "lifestyle_questionnaire_1_timestamp", 
                                                  "lifestyle_questionnaire_1_complete"))$data
LSQ1download <- LSQdownload
LSQ1complete$lifestyle_questionnaire_1_timestamp <- as.Date(LSQ1complete$lifestyle_questionnaire_1_timestamp,
                                                            "%Y-%m-%d")
LSQ1complete <- LSQ1complete[LSQ1complete[, "lifestyle_questionnaire_1_complete"] == 2, ]
  
  
secret_token <- "token"
api_url <- "api link"
LSQ1complete2 <- redcap_read_oneshot(redcap_uri = api_url,
                                       token = secret_token,
                                       fields =  c("obs_study_id",
                                                   "lifestyle_questionnaire_1_timestamp",
                                                   "lifestyle_questionnaire_1_complete"))$data
LSQ1download2 <- LSQdownload
LSQ1complete2$lifestyle_questionnaire_1_timestamp <- as.Date(LSQ1complete2$lifestyle_questionnaire_1_timestamp, 
                                                             "%Y-%m-%d")
LSQ1complete2 <- LSQ1complete2[LSQ1complete2[, "lifestyle_questionnaire_1_complete"] == 2, ]
LSQ1complete <- rbind(LSQ1complete, LSQ1complete2)
  
  
# LSQ2   
secret_token <- "token"
api_url <- "api link"
LSQ2complete <- redcap_read_oneshot(redcap_uri = api_url,
                                      token = secret_token,
                                      fields =  c("obs_study_id", 
                                                  "lifestyle_questionnaire_2_timestamp", 
                                                  "lifestyle_questionnaire_2_complete"))$data
LSQ2download <- LSQdownload
LSQ2complete$lifestyle_questionnaire_2_timestamp <- as.Date(LSQ2complete$lifestyle_questionnaire_2_timestamp,
                                                            "%Y-%m-%d")
LSQ2complete <- LSQ2complete[LSQ2complete[, "lifestyle_questionnaire_2_complete"] == 2, ]
  

secret_token <- "token"
api_url <- "api link"
LSQ2complete2 <- redcap_read_oneshot(redcap_uri = api_url,
                                       token = secret_token,
                                       fields =  c("obs_study_id",
                                                   "lifestyle_questionnaire_2_timestamp",
                                                   "lifestyle_questionnaire_2_complete"))$data
LSQ2download2 <- LSQdownload
LSQ2complete2$lifestyle_questionnaire_2_timestamp <- as.Date(LSQ2complete2$lifestyle_questionnaire_2_timestamp, 
                                                               "%Y-%m-%d")
LSQ2complete2 <- LSQ2complete2[LSQ2complete2[, "lifestyle_questionnaire_2_complete"] == 2, ]
LSQ2complete <- rbind(LSQ2complete, LSQ2complete2)
  
  
# LSQ3
secret_token <- "token"
api_url <- "api link"
LSQ3complete <- redcap_read_oneshot(redcap_uri = api_url,
                                      token = secret_token,
                                      fields =  c("obs_study_id", 
                                                  "lifestyle_questionnaire_3_timestamp", 
                                                  "lifestyle_questionnaire_3_complete"))$data
LSQ3complete <- LSQ3complete[LSQ3complete$obs_study_id < 91600000, ] #because I added SMH to MSH survey; eventually need to delete when cleaned up
LSQ3complete$lifestyle_questionnaire_3_timestamp <- as.Date(LSQ3complete$lifestyle_questionnaire_3_timestamp, 
                                                              "%Y-%m-%d")
LSQ3download <- LSQdownload
LSQ3complete <- LSQ3complete[LSQ3complete[, "lifestyle_questionnaire_3_complete"] == 2, ]
  

secret_token <- "token"
api_url <- "api link"
LSQ3complete2 <- redcap_read_oneshot(redcap_uri = api_url,
                                       token = secret_token,
                                       fields =  c("obs_study_id",
                                                   "lifestyle_questionnaire_3_timestamp",
                                                   "lifestyle_questionnaire_3_complete"))$data
  
LSQ3complete2 <- LSQ3complete2[LSQ3complete2$obs_study_id < 91600000, ] 
LSQ3complete2$lifestyle_questionnaire_3_timestamp <- as.Date(LSQ3complete2$lifestyle_questionnaire_3_timestamp, 
                                                             "%Y-%m-%d")
LSQ3download2 <- LSQdownload
LSQ3complete2 <- LSQ3complete2[LSQ3complete2[, "lifestyle_questionnaire_3_complete"] == 2, ]
LSQ3complete <- rbind(LSQ3complete, LSQ3complete2)
  
  if(LSQ1download == 200 & LSQ2download == 200 & LSQ3download == 200 & LSQ1download2 == 200 & LSQ2download2 == 200 & LSQ3download2 == 200){
    break
  } else {
    Sys.sleep(1800)
  }
}

## Completed LSQs, format change so things work with access ========================================

completed_1 <- LSQ1complete[LSQ1complete$obs_study_id %in% as.integer(gsub("-", "", check_if_completed_1)), c("obs_study_id", "lifestyle_questionnaire_1_timestamp")]
names(completed_1) <- c("OBSID", "DateComplete")
completed_1$OBSID <- gsub("^912", "912-", as.character(completed_1$OBSID))
completed_1$emailstatus <- rep("LSQ(1)Returned", nrow(completed_1))
rownames(completed_1 ) <- NULL


completed_2 <- LSQ2complete[LSQ2complete$obs_study_id %in% as.integer(gsub("-", "", check_if_completed_2)), c("obs_study_id", "lifestyle_questionnaire_2_timestamp")]
names(completed_2) <- c("OBSID", "DateComplete")
completed_2$OBSID <- gsub("^912", "912-", as.character(completed_2$OBSID))
completed_2$emailstatus <- rep("LSQ(2)Returned", nrow(completed_2))
rownames(completed_2 ) <- NULL



completed_3 <- LSQ3complete[LSQ3complete$obs_study_id %in% as.integer(gsub("-", "", check_if_completed_3)), c("obs_study_id", "lifestyle_questionnaire_3_timestamp")]
names(completed_3) <- c("OBSID", "DateComplete")
completed_3$OBSID <- gsub("^912", "912-", as.character(completed_3$OBSID))
completed_3$emailstatus <- rep("LSQ(3)Returned", nrow(completed_3))
rownames(completed_3 ) <- NULL

## LSQ completion entries in database ==============================================================


myconnection <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};
                                  DBQ= filepath")
enrolment <- sqlFetch(myconnection, "OBS Enrolment Log")


if(nrow(completed_1) > 0){
  invisible(sapply(completed_1$OBSID, function(obsid){
    PatientID <- followup[which(followup$OBSEnrolmentID == obsid), ][1, "PatientID"]
    OBSEnrolmentID <- obsid
    PatientFirstName <- followup[which(followup$OBSEnrolmentID == obsid), ][1, "PatientFirstName"]
    PatientSurname <- followup[which(followup$OBSEnrolmentID == obsid), ][1, "PatientSurname"]
    PatientSurname <- gsub("'", "''", PatientSurname)
    DateComplete <- completed_1[which(completed_1$OBSID == obsid), ][1, "DateComplete"]
    emailstatus <- "LSQ(1)Returned"
    sqlQuery(myconnection, paste0("INSERT INTO OBSFollowupLog (PatientID,",
                                  " OBSEnrolmentID, PatientFirstName,",
                                  "  PatientSurname, OBSVisitDate, \"",
                                  emailstatus,"\") VALUES ('", PatientID,"', '",
                                  OBSEnrolmentID,"', '", PatientFirstName,"', '",
                                  PatientSurname,"', '", DateComplete,
                                  "', 1)"))
  }))
}


if(nrow(completed_2) > 0){
  invisible(sapply(completed_2$OBSID, function(obsid){
    PatientID <- followup[which(followup$OBSEnrolmentID == obsid), ][1, "PatientID"]
    OBSEnrolmentID <- obsid
    PatientFirstName <- followup[which(followup$OBSEnrolmentID == obsid), ][1, "PatientFirstName"]
    PatientSurname <- followup[which(followup$OBSEnrolmentID == obsid), ][1, "PatientSurname"]
    PatientSurname <- gsub("'", "''", PatientSurname)
    DateComplete <- completed_2[which(completed_2$OBSID == obsid), ][1, "DateComplete"]
    emailstatus <- "LSQ(2)Returned"
    sqlQuery(myconnection, paste0("INSERT INTO OBSFollowupLog (PatientID,",
                                  " OBSEnrolmentID, PatientFirstName,",
                                  "  PatientSurname, OBSVisitDate, \"",
                                  emailstatus,"\") VALUES ('", PatientID,"', '",
                                  OBSEnrolmentID,"', '", PatientFirstName,"', '",
                                  PatientSurname,"', '", DateComplete,
                                  "', 1)"))
  }))
}


if(nrow(completed_3) > 0){
  invisible(sapply(completed_3$OBSID, function(obsid){
    PatientID <- followup[which(followup$OBSEnrolmentID == obsid), ][1, "PatientID"]
    OBSEnrolmentID <- obsid
    PatientFirstName <- followup[which(followup$OBSEnrolmentID == obsid), ][1, "PatientFirstName"]
    PatientSurname <- followup[which(followup$OBSEnrolmentID == obsid), ][1, "PatientSurname"]
    PatientSurname <- gsub("'", "''", PatientSurname)
    DateComplete <- completed_3[which(completed_3$OBSID == obsid), ][1, "DateComplete"]
    emailstatus <- "LSQ(3)Returned"
    sqlQuery(myconnection, paste0("INSERT INTO OBSFollowupLog (PatientID,",
                                  " OBSEnrolmentID, PatientFirstName,",
                                  "  PatientSurname, OBSVisitDate, \"",
                                  emailstatus,"\") VALUES ('", PatientID,"', '",
                                  OBSEnrolmentID,"', '", PatientFirstName,"', '",
                                  PatientSurname,"', '", DateComplete,
                                  "', 1)"))
  }))
}
close(myconnection)

## Connecting to accdb to determine new patients ===================================================

myconnection <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};
                                  DBQ= filepath")
enrolment <- sqlFetch(myconnection, "OBS Enrolment Log")
followup <- sqlFetch(myconnection, "OBS Followup Log")
close(myconnection)

enrolment <- enrolment[, c("OBSEnrolmentID", "EDC", "DIPLateEntry", "DIPCurEnrol")]
total <- merge(enrolment, followup, by="OBSEnrolmentID")

total$OBSVisitDate <- as.Date(total$OBSVisitDate)
total$DeliveryDate <- as.Date(total$DeliveryDate)
total$EDC <- as.Date(total$EDC)

#drop all subjects who have delivered more than 14 weeks ago
total <- total[total$EDC >= as.Date("05/30/16", "%m/%d/%y"), ]
total <- total[!is.na(total$OBSEnrolmentID), ]

OBSID <- as.character(unique(total$OBSEnrolmentID))

#function to see if patient gave consent or didn't withdrawl
invisible(sapply(OBSID, function(OBSEnrolmentID){
  enrol_table <- total[total$OBSEnrolmentID == OBSEnrolmentID, ]
  if(sum(enrol_table$OBSPatientWithdrawl) > 0 |
     sum(enrol_table[, "Fetal Demise/Termination"]) > 0 |
     sum(enrol_table[, "Neonatal death"]) > 0 ){ #add more variables which may eliminate subjects
    OBSID <<- OBSID[!OBSID %in% c(OBSEnrolmentID)]
  }
}))

LSQ_send_list <- data.frame(obsid = character(),
                            emailstatus = character(),
                            stringsAsFactors = FALSE)


statuses <- c("LSQ(1)Given", "LSQ(1)Followup1", "LSQ(1)Followup2", "LSQ(1)Followup3")

LSQselection <- function(enrol_table){
  for(i in 2:4){
    if(sum(enrol_table[, statuses[i]]) == 0) { #TOOK OUT LENGHT_TO_STOP argument; need to replace it, if not, delete length_to_stop
      if((enrol_table[enrol_table[, statuses[(i - 1)]] == 1, "OBSVisitDate"] + 13) <= Sys.Date()){ #has it been 14 days since the LSQ was last sent out?
        LSQ_send_list[nrow(LSQ_send_list) + 1, "obsid"] <<- enrolmentID
        LSQ_send_list[nrow(LSQ_send_list), "emailstatus"] <<- statuses[i]
        #         obsid <<- c(obsid, enrolmentID)
        #         emailstatus <<- c(emailstatus, statuses[i])
        break
      } else {
        break
      }
    }
  }
}

# creates 2 vectors (obsid, emailstatus) for people who need to be sent LSQs based on Access DB

invisible(sapply(OBSID, function(enrolmentID){
  enrol_table <- total[total$OBSEnrolmentID == enrolmentID, ]
  enrolmentID <<- enrolmentID
  print(enrolmentID)
  if(sum(enrol_table[, "LSQ(1)Returned"], enrol_table[, "LSQ(1)Followup3"],
         enrol_table[, "LSQ1Refused"], enrol_table[, "Paper LSQ1"]) < 1) {    # add LSQ1refusal here?
    if(sum(enrol_table[, "LSQ(1)Given"]) == 0){
      if(enrol_table[1, "EDC"] -(40*7) + (11*7) <= Sys.Date ()) {
        LSQ_send_list[nrow(LSQ_send_list) + 1, "obsid"] <<- enrolmentID
        LSQ_send_list[nrow(LSQ_send_list), "emailstatus"] <<- "LSQ(1)Given"
      }
    } else if(sum(enrol_table[, "LSQ(1)Given"]) == 1){
      statuses <<- gsub("\\([2-3]\\)", "(1)", statuses)
      LSQselection(enrol_table)
    }
  } else if((sum(enrol_table[, "LSQ(2)Returned"], enrol_table[, "LSQ(2)Followup3"],
                 enrol_table[, "LSQ2Refused"], enrol_table[, "Paper LSQ2"]) < 1)) {
    if(sum(enrol_table[, "LSQ(2)Given"]) == 0){
      if((enrol_table[1, "EDC"] - (40*7) + (26*7) <= Sys.Date ()) |
         ((sum(enrol_table[, "DIPLateEntry"]) > 0) ) # check if this makes sense
         
      ){
        LSQ_send_list[nrow(LSQ_send_list) + 1, "obsid"] <<- enrolmentID
        LSQ_send_list[nrow(LSQ_send_list), "emailstatus"] <<- "LSQ(2)Given"
      }
    } else if(sum(enrol_table[, "LSQ(2)Given"]) == 1) {
      statuses <<- gsub("\\([1|3]\\)", "(2)", statuses)
      LSQselection(enrol_table)
    }
  } else if((sum(enrol_table[, "LSQ(3)Returned"], enrol_table[, "LSQ(3)Followup3"],
                 enrol_table[, "LSQ3Refused"], enrol_table[, "Paper LSQ3"]) < 1)){
    if(sum(enrol_table[, "LSQ(3)Given"]) == 0){
      if(sum(!is.na(enrol_table$DeliveryDate)) == 1){
        if(enrol_table$DeliveryDate[!is.na(enrol_table$DeliveryDate)] + (5*7) <= Sys.Date()){ #checks recorded delivery date; send after 5 weeks
          LSQ_send_list[nrow(LSQ_send_list) + 1, "obsid"] <<- enrolmentID
          LSQ_send_list[nrow(LSQ_send_list), "emailstatus"] <<- "LSQ(3)Given"
        }
      } else if(sum(!is.na(enrol_table$DeliveryDate)) == 0){ #### if no deliver date, send LSQ3 based on EDC
        if(enrol_table[1, "EDC"] + (6*7) <= Sys.Date()){
          LSQ_send_list[nrow(LSQ_send_list) + 1, "obsid"] <<- enrolmentID
          LSQ_send_list[nrow(LSQ_send_list), "emailstatus"] <<- "LSQ(3)Given"
        }
      } else {
        print(paste0(enrolmentID, " has more than one delivery date"))
      }
    } else if(sum(enrol_table[, "LSQ(3)Given"]) == 1) {
      statuses <<- gsub("\\([1-2]\\)", "(3)", statuses)
      LSQselection(enrol_table)
    }
  }
}))

## Accessing patient contact information from drive ================================================

dir.create("T:temporary folder)
file.copy("T:/Patient contact information.xlsx", 
          "T:/Patient contact information.xlsx")

# connect to "Patient contact information" to get email addresses
repeat{
  tryCatch({
    myconnection <- odbcConnectExcel2007("T:Patient contact information.xlsx")
    emails <<- sqlFetch(myconnection, "Sheet1$")
    close(myconnection)
    break
  }, warning=function(w) {
    Sys.sleep(1800)
  })
}  

unlink("Patient contact information/Temp", recursive = TRUE)

emails <- emails[, c("OBSID", "E-mail")]
names(emails) <- c("OBSID", "email")
emails <- emails[complete.cases(emails), ]
emails$OBSID <- sapply(emails$OBSID, function(x){
  toString(paste0("912-", sprintf("%05.0f", x)))
})

# Validation for emails: presence of an '@' sign 
emails <- emails[grepl("@", emails$email), ]
emails <- emails[emails$OBSID %in% as.character(LSQ_send_list$obsid), ]
emails <- merge(LSQ_send_list, emails, by.x = "obsid", by.y = "OBSID", all = TRUE)
emails <- emails[!is.na(emails$emailstatus), ]
emails$email <- gsub(" ", "", emails$email)
emails <- emails[!is.na(emails$email), ]

# check if email address has '@' and domain'
if (sum(!grepl("^[A-z0-9.]+@[A-z0-9]+\\.[A-z0-9]+", emails$email)) > 0){
  emailIssues <- emails[, "obsid"][!grepl("^[A-z0-9.]+@[A-z0-9]+\\.[A-z0-9]+", emails$email)]
  

## Email Coworker about corrections that need to be made ================================================
OutApp <- COMCreate("Outlook.Application")
outMail <- OutApp$CreateItem(0)
outMail[["To"]] <- as.character("email")
outMail[["SentOnBehalfofName"]] <- "email"
outMail[["subject"]] <- "OBS Email Issues; check contact info"
outMail[["body"]] <- paste0("Hey, these patients have errors in their email addresses on the Patient Contat Information file: \n\n",
                            as.character(emailIssues))
outMail$Send()
emails <- emails[grepl("^[A-z0-9.]+@[A-z0-9]+\\.[A-z0-9]+", emails$email), ]
}



## Email patients LSQ Emails (initial, follow up 1, follow up 2) ===================================
pwlist <- read.csv("T:/file with questionnaire links and passwords.xlsx", nrow = 10000)

emailsubjects <- function(obsid){
  OutApp <- COMCreate("Outlook.Application")
  outMail <- OutApp$CreateItem(0)
  outMail[["To"]] <- as.character(emails$email[emails$obsid == obsid])
  outMail[["SentOnBehalfofName"]] <- "studyemail@organization.ca"
  if(emails$emailstatus[emails$obsid == obsid] == "LSQ(1)Given"){
    link <- pwlist$lsq1_website[pwlist$obs_subject_id == toString(obsid)]
    outMail[["subject"]] <- "Subject Content"
    outMail[["body"]] <- paste0("Here is a link\n\n",
                                link,
                                "\n\n
Your password will be in a separate email. ")
  } else if(emails$emailstatus[emails$obsid == obsid] == "LSQ(1)Followup1"){
    link <- pwlist$lsq1_website[pwlist$obs_subject_id == toString(obsid)]
    outMail[["subject"]] <- "Follow-up 1"
    outMail[["body"]] <- paste0("Reminder:\n\n",
                                link,
                                "\n\nYour password will be in a separate email. ")
  } else if(emails$emailstatus[emails$obsid == obsid] == "LSQ(1)Followup2" | emails$emailstatus[emails$obsid == obsid] == "LSQ(1)Followup3"){
    link <- pwlist$lsq1_website[pwlist$obs_subject_id == toString(obsid)]
    outMail[["subject"]] <- "Follow-up 2"
    outMail[["body"]] <- paste0("Reminder:\n\n",
                                link,
                                "\n\nYour password will be in a separate email. ")
  } else if(emails$emailstatus[emails$obsid == obsid] == "LSQ(2)Given"){
    link <- pwlist$lsq2_website[pwlist$obs_subject_id == toString(obsid)]
    outMail[["subject"]] <- "Subject Content"
    outMail[["body"]] <- paste0("Here is a link:\n\n",
                                link,
                                "\n\nYour password will be in a separate email ")
  } else if(emails$emailstatus[emails$obsid == obsid] == "LSQ(2)Followup1"){
    link <- pwlist$lsq2_website[pwlist$obs_subject_id == toString(obsid)]
    outMail[["subject"]] <- "Folloq up 1"
    outMail[["body"]] <- paste0("Here is a link ",
                                link,
                                "\n\nYour password will follow in a separate email ")
  } else if(emails$emailstatus[emails$obsid == obsid] == "LSQ(2)Followup2" | emails$emailstatus[emails$obsid == obsid] == "LSQ(2)Followup3"){
    link <- pwlist$lsq2_website[pwlist$obs_subject_id == toString(obsid)]
    outMail[["subject"]] <- "Follow Up 2"
    outMail[["body"]] <- paste0("Here is a link",
                                link,
                                "\n\nYour password will follow in a separate email ")
  } else if(emails$emailstatus[emails$obsid == obsid] == "LSQ(3)Given"){
    link <- pwlist$lsq3_website[pwlist$obs_subject_id == toString(obsid)]
    outMail[["subject"]] <- "Subject Content "
    outMail[["body"]] <- paste0("Here is a link:\n\n",
                                link,
                                "\n\nYour password will follow in a separate email ")
  } else if(emails$emailstatus[emails$obsid == obsid] == "LSQ(3)Followup1"){
    link <- pwlist$lsq3_website[pwlist$obs_subject_id == toString(obsid)]
    outMail[["subject"]] <- "Follow up 2"
    outMail[["body"]] <- paste0("Here is a link:\n\n",
                                link,
                                "\n\nYour password will follow in a separate email ")
  } else if(emails$emailstatus[emails$obsid == obsid] == "LSQ(3)Followup2" | emails$emailstatus[emails$obsid == obsid] == "LSQ(3)Followup3"){
    link <- pwlist$lsq3_website[pwlist$obs_subject_id == toString(obsid)]
    outMail[["subject"]] <- "Follow up 2"
    outMail[["body"]] <- paste0("Here is a link\n\n",
                                link,
                                "\n\nYour password will follow in a separate email.")
  }
  outMail$Send()
  Sys.sleep(5)
}

## Email subjects passwords ====================================================================
invisible(sapply(emails$obsid, emailsubjects))

emailsubjectspw <- function(obsid){
  OutApp <- COMCreate("Outlook.Application")
  outMail <- OutApp$CreateItem(0)
  outMail[["To"]] <- as.character(emails$email[emails$obsid == obsid])
  outMail[["SentOnBehalfofName"]] <- "studyemail@organization.ca"
  if(emails$emailstatus[emails$obsid == obsid] == "LSQ(1)Given" |
     emails$emailstatus[emails$obsid == obsid] == "LSQ(1)Followup1" |
     emails$emailstatus[emails$obsid == obsid] == "LSQ(1)Followup2" |
     emails$emailstatus[emails$obsid == obsid] == "LSQ(1)Followup3"){
    pw <- pwlist$lsq1_password[pwlist$obs_subject_id == toString(obsid)]
    outMail[["subject"]] <- "Password"
    outMail[["body"]] <- paste0("Here is your password:\n\n",
                                pw,
                                "\n\nPlease click the link provided in the previous email and enter the above password to access your personal Lifestyle Questionnaire.")
  } else if(emails$emailstatus[emails$obsid == obsid] == "LSQ(2)Given" |
            emails$emailstatus[emails$obsid == obsid] == "LSQ(2)Followup1" |
            emails$emailstatus[emails$obsid == obsid] == "LSQ(2)Followup2" |
            emails$emailstatus[emails$obsid == obsid] == "LSQ(2)Followup3"){
    pw <- pwlist$lsq2_password[pwlist$obs_subject_id == toString(obsid)]
    outMail[["subject"]] <- "Password"
    outMail[["body"]] <- paste0("Here is your password:\n\n",
                                pw,
                                "\n\nPlease click the link provided in the previous email and enter the above password to access your personal Lifestyle Questionnaire.")
  } else if(emails$emailstatus[emails$obsid == obsid] == "LSQ(3)Given" |
            emails$emailstatus[emails$obsid == obsid] == "LSQ(3)Followup1" |
            emails$emailstatus[emails$obsid == obsid] == "LSQ(3)Followup2" |
            emails$emailstatus[emails$obsid == obsid] == "LSQ(3)Followup3"){
    pw <- pwlist$lsq3_password[pwlist$obs_subject_id == toString(obsid)]
    outMail[["subject"]] <- "Password"
    outMail[["body"]] <- paste0("Here is your password:\n\n",
                                pw,
                                "\n\nPlease click the link provided in the previous email and enter the above password to access your personal Lifestyle Questionnaire.")
  }
  outMail$Send()
  Sys.sleep(5)
}
invisible(sapply(emails$obsid, emailsubjectspw))


## Updating accdb with information on who has been emailed or not ==================================

myconnection <- odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};
                                  DBQ= filepath")

updateaccess <- function(obsid){
  PatientID <- followup[which(followup$OBSEnrolmentID == obsid), ][1, "PatientID"]
  OBSEnrolmentID <- obsid
  PatientFirstName <- followup[which(followup$OBSEnrolmentID == obsid), ][1, "PatientFirstName"]
  PatientSurname <- followup[which(followup$OBSEnrolmentID == obsid), ][1, "PatientSurname"]
  PatientSurname <- gsub("'", "''", PatientSurname) # to compensate for people with apostrophes in their last name
  emailstatus <- emails[which(emails$obsid == obsid), ][1, "emailstatus"]
  sqlQuery(myconnection, paste0("INSERT INTO OBSFollowupLog (PatientID,",
                                " OBSEnrolmentID, PatientFirstName,",
                                "  PatientSurname, OBSVisitDate, \"",
                                emailstatus,"\") VALUES ('", PatientID,"', '",
                                OBSEnrolmentID,"', '", PatientFirstName,"', '",
                                PatientSurname,"', '", as.character(Sys.Date()),
                                "', 1)"))
}

invisible(sapply(as.character(emails$obsid), updateaccess))
close(myconnection)



## Email Coworker + myself a success message and completions ============================================================

OutApp <- COMCreate("Outlook.Application")
outMail <- OutApp$CreateItem(0)
outMail[["To"]] <- as.character("justin@organization.ca; coworker@organization.ca)
outMail[["SentOnBehalfofName"]] <- "studyemail@organization.ca"
outMail[["subject"]] <- paste0(as.character(Sys.Date()), " LSQ Distribution")
outMail[["body"]] <- paste0("Total Number of LSQs sent for the week of ",
                            as.character(Sys.Date()), ": ", nrow(emails),
                            "\n\nLSQ 1 Completed for the week: ", length(completed_1$OBSID), 
                            "\n\nLSQ 2 completed for the week: ", length(completed_2$OBSID), 
                            "\n\nLSQ 3 completed for the week: ", length(completed_3$OBSID))
outMail$Send()
outMail$Send()

## Checking LSQ2 + LSQ3 completion for EPDS analysis ======================================================
repeat{
  secret_token <- "F5F5DF07561C69F942A351DF702A25DF"
  api_url <- "https://redcap.smh.ca/redcap/api/"
  LSQ2complete <- redcap_read_oneshot(redcap_uri = api_url, 
                                      token = secret_token,
                                      fields =  c("obs_study_id", 
                                                  "lifestyle_questionnaire_2_timestamp", 
                                                  "lifestyle_questionnaire_2_complete", 
                                                  "lwk_funny",
                                                  "lwk_lookfo",
                                                  "lwk_blame",
                                                  "lwk_anxio",
                                                  "lwk_scare",
                                                  "lwk_top",
                                                  "lwk_sleep",
                                                  "lwk_miser",
                                                  "lwk_cryin",
                                                  "lwk_harm"))$data
  LSQ2download <- LSQdownload
  LSQ2complete$lifestyle_questionnaire_2_timestamp <- as.Date(LSQ2complete$lifestyle_questionnaire_2_timestamp, "%Y-%m-%d")
  
  LSQ2complete <- LSQ2complete[LSQ2complete[, "lifestyle_questionnaire_2_complete"] == 2, ]
  
  secret_token <- "81B266259963503578E01CA69255E8A8"
  api_url <- "https://redcap.smh.ca/redcap/api/"
  LSQ2complete2 <- redcap_read_oneshot(redcap_uri = api_url, 
                                       token = secret_token,
                                       fields =  c("obs_study_id", 
                                                   "lifestyle_questionnaire_2_timestamp", 
                                                   "lifestyle_questionnaire_2_complete", 
                                                   "lwk_funny",
                                                   "lwk_lookfo",
                                                   "lwk_blame",
                                                   "lwk_anxio",
                                                   "lwk_scare",
                                                   "lwk_top",
                                                   "lwk_sleep",
                                                   "lwk_miser",
                                                   "lwk_cryin",
                                                   "lwk_harm"))$data
  LSQ2download2 <- LSQdownload
  LSQ2complete2$lifestyle_questionnaire_2_timestamp <- as.Date(LSQ2complete2$lifestyle_questionnaire_2_timestamp, "%Y-%m-%d")
  LSQ2complete2 <- LSQ2complete2[LSQ2complete2[, "lifestyle_questionnaire_2_complete"] == 2, ]
  
  LSQ2complete <- rbind(LSQ2complete, LSQ2complete2)
  
  secret_token <- "5A8058DBF425BDD1284B504B893AA711"
  api_url <- "https://redcap.smh.ca/redcap/api/"
  LSQ3complete <- redcap_read_oneshot(redcap_uri = api_url, 
                                      token = secret_token,
                                      fields =  c("obs_study_id", 
                                                  "lifestyle_questionnaire_3_timestamp", 
                                                  "lifestyle_questionnaire_3_complete", 
                                                  "lweek_laugh",
                                                  "lweek_enjoy",
                                                  "lweek_blame",
                                                  "lweek_anxious",
                                                  "lweek_panic",
                                                  "lweek_top",
                                                  "lweek_unhappy",
                                                  "lweek_miserable",
                                                  "lweek_crying",
                                                  "lweek_harming"))$data
  
  LSQ3complete <- LSQ3complete[LSQ3complete$obs_study_id < 91600000, ] #because I added SMH to MSH survey; eventually need to delete when cleaned up
  
  LSQ3complete$lifestyle_questionnaire_3_timestamp <- as.Date(LSQ3complete$lifestyle_questionnaire_3_timestamp, "%Y-%m-%d")
  LSQ3download <- LSQdownload
  
  LSQ3complete <- LSQ3complete[LSQ3complete[, "lifestyle_questionnaire_3_complete"] == 2, ]

  secret_token <- "637FBFC0C50DCC153C37AE7942481534"
  api_url <- "https://redcap.smh.ca/redcap/api/"

  LSQ3complete2 <- redcap_read_oneshot(redcap_uri = api_url, 
                                       token = secret_token,
                                       fields =  c("obs_study_id", 
                                                   "lifestyle_questionnaire_3_timestamp", 
                                                   "lifestyle_questionnaire_3_complete", 
                                                   "lweek_laugh",
                                                   "lweek_enjoy",
                                                   "lweek_blame",
                                                   "lweek_anxious",
                                                   "lweek_panic",
                                                   "lweek_top",
                                                   "lweek_unhappy",
                                                   "lweek_miserable",
                                                   "lweek_crying",
                                                   "lweek_harming"))$data
  
  LSQ3complete2 <- LSQ3complete2[LSQ3complete2$obs_study_id < 91600000, ] #because I added SMH to MSH survey; eventually need to delete when cleaned up
  
  LSQ3complete2$lifestyle_questionnaire_3_timestamp <- as.Date(LSQ3complete2$lifestyle_questionnaire_3_timestamp, "%Y-%m-%d")
  LSQ3download2 <- LSQdownload
  
  LSQ3complete2 <- LSQ3complete2[LSQ3complete2[, "lifestyle_questionnaire_3_complete"] == 2, ]
  
  LSQ3complete <- rbind(LSQ3complete, LSQ3complete2)
  
  
  if(LSQ2download == 200 & LSQ3download == 200 & LSQ2download2 == 200 & LSQ3download2 == 200){
    break
  } else {
    Sys.sleep(1800)
  }
}


# subset only those who have completed since last check

LSQ2complete <- LSQ2complete[LSQ2complete$obs_study_id %in% as.numeric(gsub("-", "", completed_2$OBSID)), ]
LSQ3complete <- LSQ3complete[LSQ3complete$obs_study_id %in% as.numeric(gsub("-", "", completed_3$OBSID)), ]

## EDPS Conversion scales ==========================================================================

# EPDS conversion, LSQ2

LSQ2complete$lwk_funny <- LSQ2complete$lwk_funny - 1
LSQ2complete$lwk_lookfo <- LSQ2complete$lwk_lookfo - 1


LSQ2complete$lwk_blame[LSQ2complete$lwk_blame == 2] <- "Yes, most of the time"
LSQ2complete$lwk_blame[LSQ2complete$lwk_blame == 3] <- "Yes, some of the time"
LSQ2complete$lwk_blame[LSQ2complete$lwk_blame == 1] <- "Not very often"
LSQ2complete$lwk_blame[LSQ2complete$lwk_blame == 4] <- "No, never"

LSQ2complete$lwk_blame[LSQ2complete$lwk_blame == "Yes, most of the time"] <- 3
LSQ2complete$lwk_blame[LSQ2complete$lwk_blame == "Yes, some of the time"] <- 2
LSQ2complete$lwk_blame[LSQ2complete$lwk_blame == "Not very often"] <- 1
LSQ2complete$lwk_blame[LSQ2complete$lwk_blame == "No, never"] <- 0
LSQ2complete$lwk_blame <- as.numeric(LSQ2complete$lwk_blame)


LSQ2complete$lwk_anxio[LSQ2complete$lwk_anxio == 2] <- "No, not at all"
LSQ2complete$lwk_anxio[LSQ2complete$lwk_anxio == 1] <- "Hardly ever"
LSQ2complete$lwk_anxio[LSQ2complete$lwk_anxio == 3] <- "Yes, sometimes"
LSQ2complete$lwk_anxio[LSQ2complete$lwk_anxio == 4] <- "Yes, very often"

LSQ2complete$lwk_anxio[LSQ2complete$lwk_anxio == "No, not at all"] <- 0
LSQ2complete$lwk_anxio[LSQ2complete$lwk_anxio == "Hardly ever"] <- 1
LSQ2complete$lwk_anxio[LSQ2complete$lwk_anxio == "Yes, sometimes"] <- 2
LSQ2complete$lwk_anxio[LSQ2complete$lwk_anxio == "Yes, very often"] <- 3
LSQ2complete$lwk_anxio <- as.numeric(LSQ2complete$lwk_anxio)


LSQ2complete$lwk_scare[LSQ2complete$lwk_scare == 1] <- "Yes, quite a lot"
LSQ2complete$lwk_scare[LSQ2complete$lwk_scare == 2] <- "Yes, sometimes"
LSQ2complete$lwk_scare[LSQ2complete$lwk_scare == 3] <- "No, not much"
LSQ2complete$lwk_scare[LSQ2complete$lwk_scare == 4] <- "No, not at all"

LSQ2complete$lwk_scare[LSQ2complete$lwk_scare == "Yes, quite a lot"] <- 3
LSQ2complete$lwk_scare[LSQ2complete$lwk_scare == "Yes, sometimes"] <- 2
LSQ2complete$lwk_scare[LSQ2complete$lwk_scare == "No, not much"] <- 1
LSQ2complete$lwk_scare[LSQ2complete$lwk_scare == "No, not at all"] <- 0
LSQ2complete$lwk_scare <- as.numeric(LSQ2complete$lwk_scare)


LSQ2complete$lwk_top[LSQ2complete$lwk_top == 1] <- "Yes, most of the time I haven't been able to cope at all"
LSQ2complete$lwk_top[LSQ2complete$lwk_top == 2] <- "Yes, sometimes I haven't been coping as well as usual"
LSQ2complete$lwk_top[LSQ2complete$lwk_top == 3] <- "No, most of the time I have copied quite well"
LSQ2complete$lwk_top[LSQ2complete$lwk_top == 4] <- "No, I have been coping as well as ever"

LSQ2complete$lwk_top[LSQ2complete$lwk_top == "Yes, most of the time I haven't been able to cope at all"] <- 3
LSQ2complete$lwk_top[LSQ2complete$lwk_top == "Yes, sometimes I haven't been coping as well as usual"] <- 2
LSQ2complete$lwk_top[LSQ2complete$lwk_top == "No, most of the time I have copied quite well"] <- 1
LSQ2complete$lwk_top[LSQ2complete$lwk_top == "No, I have been coping as well as ever"] <- 0
LSQ2complete$lwk_top <- as.numeric(LSQ2complete$lwk_top)


LSQ2complete$lwk_sleep[LSQ2complete$lwk_sleep == 1] <- "Yes, most of the time"
LSQ2complete$lwk_sleep[LSQ2complete$lwk_sleep == 2] <- "Yes, sometimes"
LSQ2complete$lwk_sleep[LSQ2complete$lwk_sleep == 3] <- "Not very often"
LSQ2complete$lwk_sleep[LSQ2complete$lwk_sleep == 4] <- "No, not at all"

LSQ2complete$lwk_sleep[LSQ2complete$lwk_sleep == "Yes, most of the time"] <- 3
LSQ2complete$lwk_sleep[LSQ2complete$lwk_sleep == "Yes, sometimes"] <- 2
LSQ2complete$lwk_sleep[LSQ2complete$lwk_sleep == "Not very often"] <- 1
LSQ2complete$lwk_sleep[LSQ2complete$lwk_sleep == "No, not at all"] <- 0
LSQ2complete$lwk_sleep <- as.numeric(LSQ2complete$lwk_sleep)


LSQ2complete$lwk_miser[LSQ2complete$lwk_miser == 2] <- "Yes, most of the time"
LSQ2complete$lwk_miser[LSQ2complete$lwk_miser == 3] <- "Yes, quite often"
LSQ2complete$lwk_miser[LSQ2complete$lwk_miser == 1] <- "Not very often"
LSQ2complete$lwk_miser[LSQ2complete$lwk_miser == 4] <- "No, not at all"

LSQ2complete$lwk_miser[LSQ2complete$lwk_miser == "Yes, most of the time"] <- 3
LSQ2complete$lwk_miser[LSQ2complete$lwk_miser == "Yes, quite often"] <- 2
LSQ2complete$lwk_miser[LSQ2complete$lwk_miser == "Not very often"] <- 1
LSQ2complete$lwk_miser[LSQ2complete$lwk_miser == "No, not at all"] <- 0
LSQ2complete$lwk_miser <- as.numeric(LSQ2complete$lwk_miser)


LSQ2complete$lwk_cryin[LSQ2complete$lwk_cryin == 1] <- "Yes, most of the time"
LSQ2complete$lwk_cryin[LSQ2complete$lwk_cryin == 2] <- "Yes, quite often"
LSQ2complete$lwk_cryin[LSQ2complete$lwk_cryin == 3] <- "Only occasionally"
LSQ2complete$lwk_cryin[LSQ2complete$lwk_cryin == 4] <- "No, never"


LSQ2complete$lwk_cryin[LSQ2complete$lwk_cryin == "Yes, most of the time"] <- 3
LSQ2complete$lwk_cryin[LSQ2complete$lwk_cryin == "Yes, quite often"] <- 2
LSQ2complete$lwk_cryin[LSQ2complete$lwk_cryin == "Only occasionally"] <- 1
LSQ2complete$lwk_cryin[LSQ2complete$lwk_cryin == "No, never"] <- 0
LSQ2complete$lwk_cryin <- as.numeric(LSQ2complete$lwk_cryin)


LSQ2complete$lwk_harm[LSQ2complete$lwk_harm == 1] <- "Yes, quite often"
LSQ2complete$lwk_harm[LSQ2complete$lwk_harm == 2] <- "Sometimes"
LSQ2complete$lwk_harm[LSQ2complete$lwk_harm == 3] <- "Hardly ever"
LSQ2complete$lwk_harm[LSQ2complete$lwk_harm == 4] <- "Never"

LSQ2complete$lwk_harm[LSQ2complete$lwk_harm == "Yes, quite often"] <- 3
LSQ2complete$lwk_harm[LSQ2complete$lwk_harm == "Sometimes"] <- 2
LSQ2complete$lwk_harm[LSQ2complete$lwk_harm == "Hardly ever"] <- 1 
LSQ2complete$lwk_harm[LSQ2complete$lwk_harm == "Never"] <- 0
LSQ2complete$lwk_harm <- as.numeric(LSQ2complete$lwk_harm)



LSQ2complete$EPDS <- rowSums(LSQ2complete[, c("lwk_funny", 
                                              "lwk_lookfo",
                                              "lwk_blame",
                                              "lwk_anxio",
                                              "lwk_scare",
                                              "lwk_top",
                                              "lwk_sleep",
                                              "lwk_miser",
                                              "lwk_cryin",
                                              "lwk_harm")], na.rm = TRUE)

LSQ2EPDSreview <- LSQ2complete$obs_study_id[LSQ2complete$EPDS >= 13 | 
                                              LSQ2complete$lwk_harm == 3 | 
                                              LSQ2complete$lwk_harm == 2 | 
                                              LSQ2complete$lwk_harm == 1]

LSQ2EPDSreview <- LSQ2EPDSreview[!is.na(LSQ2EPDSreview)]
LSQ2EPDSreviewscore <- LSQ2complete$EPDS[LSQ2complete$obs_study_id %in% LSQ2EPDSreview]
LSQ2EPDSreviewharm <- LSQ2complete$lwk_harm[LSQ2complete$obs_study_id %in% LSQ2EPDSreview]


# EPDS conversion, LSQ3

LSQ3complete$lweek_laugh[LSQ3complete$lweek_laugh == 1] <- "As much as I always could"
LSQ3complete$lweek_laugh[LSQ3complete$lweek_laugh == 2] <- "Not quite so much now"
LSQ3complete$lweek_laugh[LSQ3complete$lweek_laugh == 3] <- "Definitely not so much now"
LSQ3complete$lweek_laugh[LSQ3complete$lweek_laugh == 4] <- "Not at all"

LSQ3complete$lweek_laugh[LSQ3complete$lweek_laugh == "As much as I always could"] <- 0
LSQ3complete$lweek_laugh[LSQ3complete$lweek_laugh == "Not quite so much now"] <- 1
LSQ3complete$lweek_laugh[LSQ3complete$lweek_laugh == "Definitely not so much now"] <- 2
LSQ3complete$lweek_laugh[LSQ3complete$lweek_laugh == "Not at all"] <- 3
LSQ3complete$lweek_laugh <- as.numeric(LSQ3complete$lweek_laugh)


LSQ3complete$lweek_enjoy[LSQ3complete$lweek_enjoy == 1] <- "As much as I ever did"
LSQ3complete$lweek_enjoy[LSQ3complete$lweek_enjoy == 2] <- "Rather less than I used to"
LSQ3complete$lweek_enjoy[LSQ3complete$lweek_enjoy == 3] <- "Definitely less than I used to"
LSQ3complete$lweek_enjoy[LSQ3complete$lweek_enjoy == 4] <- "Hardly at all"

LSQ3complete$lweek_enjoy[LSQ3complete$lweek_enjoy == "As much as I ever did"] <- 0
LSQ3complete$lweek_enjoy[LSQ3complete$lweek_enjoy == "Rather less than I used to"] <- 1
LSQ3complete$lweek_enjoy[LSQ3complete$lweek_enjoy == "Definitely less than I used to"] <- 2
LSQ3complete$lweek_enjoy[LSQ3complete$lweek_enjoy == "Hardly at all"] <- 3
LSQ3complete$lweek_enjoy <- as.numeric(LSQ3complete$lweek_enjoy)


LSQ3complete$lweek_blame[LSQ3complete$lweek_blame == 2] <- "Yes, most of the time"
LSQ3complete$lweek_blame[LSQ3complete$lweek_blame == 3] <- "Yes, some of the time"
LSQ3complete$lweek_blame[LSQ3complete$lweek_blame == 1] <- "Not very often"
LSQ3complete$lweek_blame[LSQ3complete$lweek_blame == 4] <- "No, never"

LSQ3complete$lweek_blame[LSQ3complete$lweek_blame == "Yes, most of the time"] <- 3
LSQ3complete$lweek_blame[LSQ3complete$lweek_blame == "Yes, some of the time"] <- 2
LSQ3complete$lweek_blame[LSQ3complete$lweek_blame == "Not very often"] <- 1
LSQ3complete$lweek_blame[LSQ3complete$lweek_blame == "No, never"] <- 0
LSQ3complete$lweek_blame <- as.numeric(LSQ3complete$lweek_blame)


LSQ3complete$lweek_anxious[LSQ3complete$lweek_anxious == 1] <- "No, not at all"
LSQ3complete$lweek_anxious[LSQ3complete$lweek_anxious == 2] <- "Hardly ever"
LSQ3complete$lweek_anxious[LSQ3complete$lweek_anxious == 3] <- "Yes, sometimes"
LSQ3complete$lweek_anxious[LSQ3complete$lweek_anxious == 4] <- "Yes, very often"

LSQ3complete$lweek_anxious[LSQ3complete$lweek_anxious == "No, not at all"] <- 0
LSQ3complete$lweek_anxious[LSQ3complete$lweek_anxious == "Hardly ever"] <- 1
LSQ3complete$lweek_anxious[LSQ3complete$lweek_anxious == "Yes, sometimes"] <- 2
LSQ3complete$lweek_anxious[LSQ3complete$lweek_anxious == "Yes, very often"] <- 3
LSQ3complete$lweek_anxious <- as.numeric(LSQ3complete$lweek_anxious)


LSQ3complete$lweek_panic[LSQ3complete$lweek_panic == 1] <- "Yes, quite a lot"
LSQ3complete$lweek_panic[LSQ3complete$lweek_panic == 2] <- "Yes, sometimes"
LSQ3complete$lweek_panic[LSQ3complete$lweek_panic == 3] <- "No, not much"
LSQ3complete$lweek_panic[LSQ3complete$lweek_panic == 4] <- "No, not at all"

LSQ3complete$lweek_panic[LSQ3complete$lweek_panic == "Yes, quite a lot"] <- 3
LSQ3complete$lweek_panic[LSQ3complete$lweek_panic == "Yes, sometimes"] <- 2
LSQ3complete$lweek_panic[LSQ3complete$lweek_panic == "No, not much"] <- 1
LSQ3complete$lweek_panic[LSQ3complete$lweek_panic == "No, not at all"] <- 0
LSQ3complete$lweek_panic <- as.numeric(LSQ3complete$lweek_panic)


LSQ3complete$lweek_top[LSQ3complete$lweek_top == 1] <- "Yes, most of the time I haven't been able to cope at all"
LSQ3complete$lweek_top[LSQ3complete$lweek_top == 2] <- "Yes, sometimes I haven't been coping as well as usual"
LSQ3complete$lweek_top[LSQ3complete$lweek_top == 3] <- "No, most of the time I have coped quite well"
LSQ3complete$lweek_top[LSQ3complete$lweek_top == 4] <- "No, I have been coping as well as ever"

LSQ3complete$lweek_top[LSQ3complete$lweek_top == "Yes, most of the time I haven't been able to cope at all"] <- 3
LSQ3complete$lweek_top[LSQ3complete$lweek_top == "Yes, sometimes I haven't been coping as well as usual"] <- 2
LSQ3complete$lweek_top[LSQ3complete$lweek_top == "No, most of the time I have coped quite well"] <- 1
LSQ3complete$lweek_top[LSQ3complete$lweek_top == "No, I have been coping as well as ever"] <- 0
LSQ3complete$lweek_top <- as.numeric(LSQ3complete$lweek_top)


LSQ3complete$lweek_unhappy[LSQ3complete$lweek_unhappy == 1] <- "Yes, most of the time"
LSQ3complete$lweek_unhappy[LSQ3complete$lweek_unhappy == 2] <- "Yes, sometimes"
LSQ3complete$lweek_unhappy[LSQ3complete$lweek_unhappy == 3] <- "Not very often"
LSQ3complete$lweek_unhappy[LSQ3complete$lweek_unhappy == 4] <- "No, not at all"

LSQ3complete$lweek_unhappy[LSQ3complete$lweek_unhappy == "Yes, most of the time"] <- 3
LSQ3complete$lweek_unhappy[LSQ3complete$lweek_unhappy == "Yes, sometimes"] <- 2
LSQ3complete$lweek_unhappy[LSQ3complete$lweek_unhappy == "Not very often"] <- 1
LSQ3complete$lweek_unhappy[LSQ3complete$lweek_unhappy == "No, not at all"] <- 0
LSQ3complete$lweek_unhappy <- as.numeric(LSQ3complete$lweek_unhappy)


LSQ3complete$lweek_miserable[LSQ3complete$lweek_miserable == 1] <- "Yes, most of the time"
LSQ3complete$lweek_miserable[LSQ3complete$lweek_miserable == 2] <- "Yes, quite often"
LSQ3complete$lweek_miserable[LSQ3complete$lweek_miserable == 3] <- "Not very often"
LSQ3complete$lweek_miserable[LSQ3complete$lweek_miserable == 4] <- "No, not at all"

LSQ3complete$lweek_miserable[LSQ3complete$lweek_miserable == "Yes, most of the time"] <- 3
LSQ3complete$lweek_miserable[LSQ3complete$lweek_miserable == "Yes, quite often"] <- 2
LSQ3complete$lweek_miserable[LSQ3complete$lweek_miserable == "Not very often"] <- 1
LSQ3complete$lweek_miserable[LSQ3complete$lweek_miserable == "No, not at all"] <- 0
LSQ3complete$lweek_miserable <- as.numeric(LSQ3complete$lweek_miserable)


LSQ3complete$lweek_crying[LSQ3complete$lweek_crying == 1] <- "Yes, most of the time"
LSQ3complete$lweek_crying[LSQ3complete$lweek_crying == 2] <- "Yes, quite often"
LSQ3complete$lweek_crying[LSQ3complete$lweek_crying == 3] <- "Only occasionally"
LSQ3complete$lweek_crying[LSQ3complete$lweek_crying == 4] <- "No, never"

LSQ3complete$lweek_crying[LSQ3complete$lweek_crying == "Yes, most of the time"] <- 3
LSQ3complete$lweek_crying[LSQ3complete$lweek_crying == "Yes, quite often"] <- 2
LSQ3complete$lweek_crying[LSQ3complete$lweek_crying == "Only occasionally"] <- 1
LSQ3complete$lweek_crying[LSQ3complete$lweek_crying == "No, never"] <- 0
LSQ3complete$lweek_crying <- as.numeric(LSQ3complete$lweek_crying)


LSQ3complete$lweek_harming[LSQ3complete$lweek_harming == 4] <- "Yes, quite often"
LSQ3complete$lweek_harming[LSQ3complete$lweek_harming == 5] <- "Sometimes"
LSQ3complete$lweek_harming[LSQ3complete$lweek_harming == 6] <- "Hardly ever"
LSQ3complete$lweek_harming[LSQ3complete$lweek_harming == 7] <- "Never"

LSQ3complete$lweek_harming[LSQ3complete$lweek_harming == "Yes, quite often"] <- 3
LSQ3complete$lweek_harming[LSQ3complete$lweek_harming == "Sometimes"] <- 2
LSQ3complete$lweek_harming[LSQ3complete$lweek_harming == "Hardly ever"] <- 1
LSQ3complete$lweek_harming[LSQ3complete$lweek_harming == "Never"] <- 0
LSQ3complete$lweek_harming <- as.numeric(LSQ3complete$lweek_harming)



LSQ3complete$EPDS <- rowSums(LSQ3complete[, c("lweek_laugh",
                                              "lweek_enjoy",
                                              "lweek_blame",
                                              "lweek_anxious",
                                              "lweek_panic",
                                              "lweek_top",
                                              "lweek_unhappy",
                                              "lweek_miserable",
                                              "lweek_crying",
                                              "lweek_harming")], na.rm = TRUE)

LSQ3EPDSreview <- LSQ3complete$obs_study_id[(LSQ3complete$EPDS >= 13 | 
                                               LSQ3complete$lweek_harming == 3 | 
                                               LSQ3complete$lweek_harming == 2 | 
                                               LSQ3complete$lweek_harming == 1) &
                                              !is.na(LSQ3complete$lweek_harming)]

LSQ3EPDSreview <- LSQ3EPDSreview[!is.na(LSQ3EPDSreview)]
LSQ3EPDSreviewscore <- LSQ3complete$EPDS[LSQ3complete$obs_study_id %in% LSQ3EPDSreview]
LSQ3EPDSreviewharm <- LSQ3complete$lweek_harming[LSQ3complete$obs_study_id %in% LSQ3EPDSreview]

if((length(LSQ2EPDSreview)+ length(LSQ3EPDSreview)) > 0){
  
  # create string
  EPDStoreview <- c(LSQ2EPDSreview, LSQ3EPDSreview)
  EPDStoreviewscore <- c(LSQ2EPDSreviewscore, LSQ3EPDSreviewscore)
  EPDStoreviewharm <- c(LSQ2EPDSreviewharm, LSQ3EPDSreviewharm)
  
  EPDStoreviewharm[EPDStoreviewharm == 0] <- "Never"
  EPDStoreviewharm[EPDStoreviewharm == 1] <- "Hardly ever"
  EPDStoreviewharm[EPDStoreviewharm == 2] <- "Sometimes"
  EPDStoreviewharm[EPDStoreviewharm == 3] <- "Yes, quite often"
  
  
  complete_insert <- as.character()
  
  for(i in 1:length(EPDStoreview)){
    complete_insert <<- as.character(paste0(complete_insert, 
                                            EPDStoreview[i], 
                                            " (EPDS Score: ",
                                            EPDStoreviewscore[i],
                                            "; 'Harming myself' answer: ",
                                            EPDStoreviewharm[i],
                                            ")\n"))
  }
  
  OutApp <- COMCreate("Outlook.Application")
  outMail <- OutApp$CreateItem(0)
  outMail[["To"]] <- as.character("coworker@organization.ca")
  outMail[["SentOnBehalfofName"]] <- "justin@organization.ca"
  outMail[["subject"]] <- "EPDS Followup List"
  outMail[["body"]] <- paste0("Hey ,\n\nFor this week the following OBS subjects need to be followed up based on their EPDS score:\n\n", 
                              complete_insert 
                             )
  
  outMail$Send()
  
} else {
  OutApp <- COMCreate("Outlook.Application")
  outMail <- OutApp$CreateItem(0)
outMail[["To"]] <- as.character("coworker@organization.ca")
  outMail[["SentOnBehalfofName"]] <- "justin@organization.ca"
  outMail[["subject"]] <- "EPDS Followup List"
  outMail[["body"]] <- paste0("Hey,\n\nThere are no EPDS followups this week. Yay!\n\")
  
  outMail$Send()
}


# End script
