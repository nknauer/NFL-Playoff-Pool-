library(shiny)
library(shinyWidgets)
library(gmailr)
library(rdrop2)
library(readxl)
library(DT)
library(lubridate)
library(gargle)

# Define UI ----
ui <- fluidPage(
  setBackgroundColor("ghostwhite"),
  titlePanel("Bill Buss Submission Collection"),
  sidebarLayout(
    sidebarPanel(
      HTML('<center><img src="bill_buss_logo.png" width="140"></center>'),
      helpText("Purpose of this tool: "),
      helpText("Collect email entries for Bill Buss Pool directly from your gmail. This tool will also backup your entries into a separate folder in your DropBox Account."),
      helpText("Use the Gmail Account: bill.buss.pool.submissions@gmail.com. This is also the email linked to your DropBox Account."),
      helpText("Instructions: "),
      helpText("Make sure all email submission attachments are in the respective folder on your Gmail - 'Submissions_PlayoffYear'. For example, if the playoffs are going to occur in 2020, then name your gmail folder: Submissions_2020 "),
      helpText("By clicking the below button, all attachments from this Gmail Folder will get downloaded, backed up in DropBox, and displayed on the right panel"),
      helpText("As all of this is taking place, you can see the progress on the bottom right of the screen for each section of the task."),
      actionButton("GMAIL_BUTTON", "Display Entries")),
    mainPanel(
        tabsetPanel(type="tabs",
                    tabPanel("All Submissions", DT::dataTableOutput("table1"))
        ))))

# Define server logic ----
server <- function(input, output, session) {

  temp_directory_for_files<-
    tempdir()  
##AUTHENTICATE GMAIL EMAIL 
observeEvent(input$GMAIL_BUTTON, {
  ##options(httr_oob_default = TRUE, httr_oauth_cache=TRUE)
  
  ##options(httr_oob_default=T)
  ##gmailr::clear_token()
  gm_auth_configure(path = "~/Gmail Attachments and DropBox Backup_v2/json/billbussapi.json")
  ##gmailr::use_secret_file("~/Gmail Attachments and DropBox Backup_v2/json/billbussapi.json")
  gm_auth(email = TRUE, cache = ".secret")
  ##gm_auth_configure(path = "credentials.json")
        ##gmailr::gm_auth()
        raw_folder_outputs<-gmailr::gm_labels(user_id = "me")
        year_of_pool<- ifelse(month(Sys.Date())>2, year(Sys.Date())+1, year(Sys.Date()))
        raw_folder_output_labels<-raw_folder_outputs$labels
        ##Extract the ID for the Gmail Folder you selected
        folder_id<-sapply(raw_folder_output_labels[sapply(raw_folder_output_labels, 
                                                          `[[`, 
                                                          "name") == paste("Submissions",year_of_pool,sep = "_")],`[[`, "id")
        
        ##Extract all the messages in that folder from the Gmail folder ID
        mssgs <- 
          gm_messages(label_ids = folder_id,
                   user_id = "me")
        
        
        ##Loading Bar
        withProgress(message = 'Downloading Email Attachments', value = 0, {
          n <- length(gmailr::gm_id(mssgs))
          
        
        ##Download all attachments from the Gmail Folder to your Adobe local folder file path
        for (i in 1:length(gmailr::gm_id(mssgs))){
          ids = gmailr::gm_id(mssgs)
          Mn = gm_message(ids[i], user_id = "me")
          path = temp_directory_for_files
          gm_save_attachments(Mn, attachment_id = NULL, path, user_id = "me")
          incProgress(1/n, detail = paste("Email", i, "of", n))
          Sys.sleep(0.1)
        }})

setwd(temp_directory_for_files)

##Call your file path to upload all Adobe files to R
require(tidyverse)

Bill_Buss_df <- list.files(path = temp_directory_for_files,
                           full.names = TRUE,
                           recursive = TRUE,
                           pattern = "*.xls") %>% 
  tibble::as_tibble() %>%
  mutate(sheetName = map(value, readxl::excel_sheets)) %>%
  unnest(sheetName) %>%
  filter(.,sheetName=="ENTRY DATA LINE") %>%
  mutate(myFiles = purrr::map2(value, sheetName, function(x,y) {
    readxl::read_excel(x, sheet = paste(y), skip = 3)})) %>% 
  unnest(myFiles)

Bill_Buss_df$value<-NULL
Bill_Buss_df$sheetName<-NULL

colnames(Bill_Buss_df)[2]<- "Wild Card 1 TM_Away"
colnames(Bill_Buss_df)[3]<- "Wild Card 1 SC_Away"
colnames(Bill_Buss_df)[4]<- "Wild Card 1 TM_Home"
colnames(Bill_Buss_df)[5]<- "Wild Card 1 SC_Home"
colnames(Bill_Buss_df)[6]<- "Wild Card 2 TM_Away"
colnames(Bill_Buss_df)[7]<- "Wild Card 2 SC_Away"
colnames(Bill_Buss_df)[8]<- "Wild Card 2 TM_Home"
colnames(Bill_Buss_df)[9]<- "Wild Card 2 SC_Home"
colnames(Bill_Buss_df)[10]<- "Wild Card 3 TM_Away"
colnames(Bill_Buss_df)[11]<- "Wild Card 3 SC_Away"
colnames(Bill_Buss_df)[12]<- "Wild Card 3 TM_Home"
colnames(Bill_Buss_df)[13]<- "Wild Card 3 SC_Home"
colnames(Bill_Buss_df)[14]<- "Wild Card 4 TM_Away"
colnames(Bill_Buss_df)[15]<- "Wild Card 4 SC_Away"
colnames(Bill_Buss_df)[16]<- "Wild Card 4 TM_Home"
colnames(Bill_Buss_df)[17]<- "Wild Card 4 SC_Home"
colnames(Bill_Buss_df)[18]<- "Wild Card 5 TM_Away"
colnames(Bill_Buss_df)[19]<- "Wild Card 5 SC_Away"
colnames(Bill_Buss_df)[20]<- "Wild Card 5 TM_Home"
colnames(Bill_Buss_df)[21]<- "Wild Card 5 SC_Home"
colnames(Bill_Buss_df)[22]<- "Wild Card 6 TM_Away"
colnames(Bill_Buss_df)[23]<- "Wild Card 6 SC_Away"
colnames(Bill_Buss_df)[24]<- "Wild Card 6 TM_Home"
colnames(Bill_Buss_df)[25]<- "Wild Card 6 SC_Home"
colnames(Bill_Buss_df)[26]<- "Div Game 1 TM_Away"
colnames(Bill_Buss_df)[27]<- "Div Game 1 SC_Away"
colnames(Bill_Buss_df)[28]<- "Div Game 1 TM_Home"
colnames(Bill_Buss_df)[29]<- "Div Game 1 SC_Home"
colnames(Bill_Buss_df)[30]<- "Div Game 2 TM_Away"
colnames(Bill_Buss_df)[31]<- "Div Game 2 SC_Away"
colnames(Bill_Buss_df)[32]<- "Div Game 2 TM_Home"
colnames(Bill_Buss_df)[33]<- "Div Game 2 SC_Home"
colnames(Bill_Buss_df)[34]<- "Div Game 3 TM_Away"
colnames(Bill_Buss_df)[35]<- "Div Game 3 SC_Away"
colnames(Bill_Buss_df)[36]<- "Div Game 3 TM_Home"
colnames(Bill_Buss_df)[37]<- "Div Game 3 SC_Home"
colnames(Bill_Buss_df)[38]<- "Div Game 4 TM_Away"
colnames(Bill_Buss_df)[39]<- "Div Game 4 SC_Away"
colnames(Bill_Buss_df)[40]<- "Div Game 4 TM_Home"
colnames(Bill_Buss_df)[41]<- "Div Game 4 SC_Home"
colnames(Bill_Buss_df)[42]<- "AFC Champ TM_Away"
colnames(Bill_Buss_df)[43]<- "AFC Champ SC_Away"
colnames(Bill_Buss_df)[44]<- "AFC Champ TM_Home"
colnames(Bill_Buss_df)[45]<- "AFC Champ SC_Home"
colnames(Bill_Buss_df)[46]<- "NFC Champ TM_Away"
colnames(Bill_Buss_df)[47]<- "NFC Champ SC_Away"
colnames(Bill_Buss_df)[48]<- "NFC Champ TM_Home"
colnames(Bill_Buss_df)[49]<- "NFC Champ SC_Home"
colnames(Bill_Buss_df)[50]<- "Superbowl TM_Away"
colnames(Bill_Buss_df)[51]<- "Superbowl SC_Away"
colnames(Bill_Buss_df)[52]<- "Superbowl TM_Home"
colnames(Bill_Buss_df)[53]<- "Superbowl SC_Home"

##setwd("~/")
##readr::write_csv(Bill_Buss_df, 'test_11_22.csv')

output$table1 <- DT::renderDataTable({
  datatable(Bill_Buss_df, extensions = 'Buttons', options = list(
    initComplete = JS(
      "function(settings, json) {",
      "$(this.api().table().header()).css({'background-color': '#0894e5', 'color': '#000000'});",
      "}"),
    pageLength = 100,
    dom = 'Bfrtip',
    buttons = 
      list('copy', 'print', list(
        extend = 'collection',
        buttons = c('csv', 'excel', 'pdf'),
        text = 'Download'
      ))
    
  )
  ) 
})

})






}

# Run the app ----
shinyApp(ui = ui, server = server)


