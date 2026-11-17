# /////////////////////////////////////////////////////////////////////////////
# Carlos Rodriguez, PhD. CU Anschutz, Dept. of Fam. Medicine
# Render & Email Medical Interpreter Report

# /////////////////////////////////////////////////////////////////////////////


#%% Libraries -----------------------------------------------------------------
library(quarto)
library(Microsoft365R)
library(stringr)

#%% Render the .qmd file ------------------------------------------------------
# Set the project directory
# proj_dir <- "2025_interpreter_training_zha_bull"

# Set the file to be rendered
qmd_file <- "D:/dfm_inquiry/2025_interpreter_training_zha_bull/scripts/medical_interpreter_report.qmd"

# Render the report
quarto_render(
  input = qmd_file
)


#%% Email the rendered file ---------------------------------------------------
# Authenticate and get your Outlook object
outlook <- get_business_outlook()  # or get_personal_outlook()

# Create an email
email <- outlook$create_email(
  body = "Automated email delivery: DFM Inquiry - Resident medical interpreter training.",
  subject = "Resident medical interpreter training report",
  to = c(
    "jessica.zha@dhha.org",
    "lauren.bull@dhha.org",
    "carlos.i.rodriguez@cuanschutz.edu")  
)

# Attach rendered file
docx_file <- str_replace(qmd_file, ".qmd", ".docx")

email$add_attachment(docx_file)

# Send email
email$send()