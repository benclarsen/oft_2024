
output_dir <- "04_Analysis/Results/women" 
project_name <- "OFT2025_women"



# Helper function to construct full output path
make_output_path <- function(filename) {
  if (!dir.exists(output_dir)) dir.create(output_dir)
  file.path(output_dir, paste0(project_name, "_", filename))
}


# Load required libraries
library(tidyverse)#
library(boot)
library(openxlsx)


## Load and prepare data -----
### Raw case data --------

caselist <- read_delim("02_Data/deid/caselist_deid.csv", delim = ",") %>% filter(sex == "women") %>% filter(when_occurred != "Other" | is.na(when_occurred)) %>% filter(subsequent_cat != "exacerbation")


### Exposure data --------

data_exposure <- read_delim("02_Data/deid/exposure_wide_deid.csv", delim = ",")
data_exposure <- data_exposure %>% 
  group_by(sex)%>%
  filter(sex == "women") %>%
  summarise(
    match = sum(match)/60,
    training = sum(training)/60,
    total = sum(total)/60
  ) %>% print()

data_exposure_days <- read_delim("02_Data/raw/exposure/exposure_days.csv", delim = ";", locale = locale(decimal_mark = ","))
data_exposure_days <- data_exposure_days %>% filter(sex =="women")
player_days <- sum(data_exposure_days$exposure_days)


### Load OSIICS v15 -------
my_locale <- locale(encoding = "UTF-8")
osiics_15 <- read_delim("02_Data/raw/osiics_15.csv", delim = ";", locale = my_locale)

# Clean whitespace from classification fields
osiics_15 <- osiics_15 %>%
  mutate(
    osiics_15_body_part = str_trim(osiics_15_body_part, side = "right"),
    osiics_15_tissue = str_trim(osiics_15_tissue, side = "right"),
    osiics_15_pathology = str_trim(osiics_15_pathology, side = "both"),
    osiics_15_diagnosis = str_trim(osiics_15_diagnosis, side = "right")
  )

### Merge OSIICS Info into Caselist ------
#caselist <- caselist %>%. #ubbecessary for this project 
# left_join(osiics_15, by = "osiics_15_code")  

### Define Factor Levels -----------
### (IOC Consensus, for table structure)

# Level 1: Body Part
level_1_order <- c(
  "Head", "Neck", "Shoulder", "Upper arm", "Elbow", "Forearm", "Wrist", "Hand",
  "Chest", "Thoracic spine", "Thoracic Spine", "Lumbosacral", "Lumbar Spine",
  "Abdomen", "Hip/groin", "Groin/hip", "Hip", "Groin", "Thigh", "Knee",
  "Lower leg", "Ankle", "Foot", "Region unspecified",
  "Single injury crossing two or more regions", "Medical", "Cardiovascular",
  "Dermatological", "Dental", "Endocrinological", "Gastrointestinal",
  "Genitourinary", "Hematologic", "Musculoskeletal", "Neurological",
  "Opthalmological", "Otological", "Psychiatric/psychological", "Respiratory",
  "Thermoregulatory", "Multiple systems", "Unknown or not specified"
)

# Level 2: Tissue Type
level_2_order <- c(
  "All", "Muscle/tendon", "Nervous", "Nervous system", "Bone",
  "Cartilage/synovium/bursa", "Ligament/joint capsule", "Superficial tissues/skin",
  "Vessels", "Stump", "Internal organs", "Non-specific", "Allergic",
  "Environmental – exercise-related", "Environmental – non-exercise",
  "Immunological/inflammatory", "Infection", "Neoplasm", "Metabolic/nutritional",
  "Thrombotic/haemorrhagic", "Degenerative or chronic condition",
  "Developmental anomaly", "Drug-related/poisoning", "Multiple",
  "Unknown, or not specified"
)

# Level 3: Pathology
level_3_order <- c(
  "All", "Muscle injury", "Muscle contusion", "Contusion/vascular",
  "Muscle compartment syndrome", "Tendinopathy", "Tendon rupture",
  "Brain & spinal cord injury", "Brain/Spinal cord injury", "Peripheral nerve injury",
  "Nerve injury", "Fracture", "Bone stress injury", "Bone contusion",
  "Avascular necrosis", "Physis injury", "Cartilage", "Cartilage injury",
  "Arthritis", "Synovitis / capsulitis", "Synovitis/capsulitis", "Bursitis",
  "Joint sprain (ligament tear or acute instability)", "Joint sprain",
  "Chronic instability", "Contusion (superficial)", "Superficial contusion",
  "Laceration", "Abrasion", "Vascular trauma", "Stump injury", "Organ trauma",
  "Injury without tissue type specified", "Pain without tissue type specified",
  "Unknown"
)

# Level 4: Diagnosis
level_4_order <- c("All", unique(osiics_15$osiics_15_diagnosis))

# Apply Factor Levels
caselist <- caselist %>%
  mutate(
    osiics_15_body_part = factor(str_trim(osiics_15_body_part, side = "right"), levels = level_1_order),
    osiics_15_tissue = factor(str_trim(osiics_15_tissue, side = "right"), levels = level_2_order),
    # Uncomment if needed:
    # osiics_15_pathology = factor(str_trim(osiics_15_pathology, side = "both"), levels = level_3_order),
    osiics_15_diagnosis = factor(str_trim(osiics_15_diagnosis, side = "right"), levels = level_4_order),
    case_id = row_number()
  ) %>%
  filter(subsequent_cat != "Exacerbations")  # Exclude exacerbations

# Optional: Check for Unmatched Values
setdiff(caselist$osiics_15_body_part, level_1_order)
setdiff(caselist$osiics_15_tissue, level_2_order)
setdiff(caselist$osiics_15_pathology, level_3_order)




# Basic descriptive statistics -------

### Number of cases by problem type and subsequent category -----
caselist %>%
  group_by(problem_type, subsequent_cat) %>%
  summarise(n_cases = n_distinct(case_id), .groups = "drop") %>%
  print()

### Number of time-loss cases by problem type ----------------
caselist %>%
  filter(timeloss_cat == "timeloss") %>%
  group_by(problem_type) %>%
  summarise(n_cases = n_distinct(case_id), .groups = "drop") %>%
  print()

### Number of involved players by problem type -----
caselist %>%
  group_by(problem_type) %>%
  summarise(n_players = n_distinct(player_id), .groups = "drop") %>%
  print()



# Incidence analyses -----

# Incidence and Severity Calculation Function
# Calculates incidence rate per 1000 hours and Poisson confidence intervals



# Notes:
# This code calculates incidence rates and confidence intervals.
# - If grouping_vars is empty (grouping_vars <- c()), it calculates overall incidence.
# - If grouping_vars contains variable names (e.g., c("body_area", "tissue")), 
#   it calculates incidence rates for each combination of these variables.
# - To change the grouping, modify the grouping_vars vector at the beginning of the script.
# - The code automatically adapts to calculate either grouped or overall incidence based on grouping_vars.



# Assumptions of this method:
# 1. Poisson distribution: Injuries follow a Poisson distribution (rare, independent events).
# 2. Constant rate: The injury rate remains constant over the exposure period.
# 3. Independence: Each injury event is independent of other injury events.
# 4. Large sample size: Sufficiently large sample for Poisson distribution approximations.
# 5. Accurate exposure time: Total exposure time (hours) is accurately measured and reported.
# 6. Complete injury reporting: All relevant injuries are correctly identified and reported.
# 7. Homogeneity: The population at risk is homogeneous in terms of injury risk.
# 8. No repeated injuries: Method doesn't account for multiple injuries to the same individual.
# 9. Linear relationship: Assumed linear relationship between exposure time and injury risk.
# 10. Rare events: Injuries are relatively rare compared to the total exposure time.

# Note: Violations of these assumptions could lead to biased or inaccurate estimates of injury rates.






calculate_incidence_severity <- function(caselist, exposure, grouping_vars = c()) {
  options(scipen = 999)  # Avoid scientific notation
  
  # Internal function to calculate Poisson confidence intervals
  calculate_poisson_ci <- function(injuries, exposure) {
    result <- poisson.test(injuries, T = exposure, conf.level = 0.95)
    data.frame(
      ci_lower = result$conf.int[1] * 1000,
      ci_upper = result$conf.int[2] * 1000
    )
  }
  
  # Apply grouping if specified
  if (length(grouping_vars) > 0) {
    caselist <- caselist %>% group_by(across(all_of(grouping_vars)))
  }
  
  caselist %>%
    summarise(
      n_cases = n(),
      total_timeloss = sum(timeloss, na.rm = TRUE),
      median_timeloss = median(timeloss, na.rm = TRUE),
      q1_timeloss = quantile(timeloss, 0.25, na.rm = TRUE),
      q3_timeloss = quantile(timeloss, 0.75, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(
      total_exposure_hours = exposure,
      incidence_rate = (n_cases / total_exposure_hours) * 1000
    ) %>%
    rowwise() %>%
    mutate(calculate_poisson_ci(n_cases, total_exposure_hours)) %>%
    ungroup() %>%
    select(-total_exposure_hours)
}


# Helper function to run and label each analysis
run_analysis <- function(caselist_input, exposure_input, outcome_label, definition_label) {
  calculate_incidence_severity(caselist_input, exposure_input) %>%
    mutate(outcome = outcome_label, definition = definition_label) %>%
    select(outcome, definition, everything()) %>%
    print()
}

# Exposure totals
total_exposure <- sum(data_exposure$total)
match_exposure <- data_exposure$match
training_exposure <- data_exposure$training

### Perform analyses ----
# Injury analyses (using hours as exposure)
all_ma <- run_analysis(caselist %>% filter(problem_type == "Injury"), total_exposure, "All injuries", "Medical attention")
all_tl <- run_analysis(caselist %>% filter(problem_type == "Injury", timeloss_cat == "timeloss"), total_exposure, "All injuries", "Time loss")

match_ma <- run_analysis(caselist %>% filter(problem_type == "Injury", when_occurred == "Match"), match_exposure, "Match injuries", "Medical attention")
match_tl <- run_analysis(caselist %>% filter(problem_type == "Injury", when_occurred == "Match", timeloss_cat == "timeloss"), match_exposure, "Match injuries", "Time loss")

tr_ma <- run_analysis(caselist %>% filter(problem_type == "Injury", when_occurred == "Training"), training_exposure, "Training injuries", "Medical attention")
tr_tl <- run_analysis(caselist %>% filter(problem_type == "Injury", when_occurred == "Training", timeloss_cat == "timeloss"), training_exposure, "Training injuries", "Time loss")

# Illness analyses (using player-days as exposure)
all_ma_illness <- run_analysis(caselist %>% filter(problem_type == "Illness"), player_days, "All illnesses", "Medical attention")
all_tl_illness <- run_analysis(caselist %>% filter(problem_type == "Illness", timeloss_cat == "timeloss"), player_days, "All illnesses", "Time loss")

# 4. Combine and Export Results

result_incidence <- bind_rows(
  all_ma, match_ma, tr_ma,
  all_tl, match_tl, tr_tl,
  all_ma_illness, all_tl_illness
)




# Burden analysis -------


# Burden calculations --------

# Note on Confidence Interval Calculation:
# This code attempts to calculate a confidence interval for any group with at least
# two cases (n > 1) that show variability in the bootstrapped samples.
# - If there's only one case or if all bootstrapped values are identical (no variability),
#   the confidence interval is set to NA.
# - For groups with n > 1 and variability in the data, a CI is calculated.
# - The code does not enforce a minimum number of cases for CI calculation beyond n > 1.
# - While CIs can be calculated for small n, they may be very wide and less informative.
# - Interpret CIs from small samples with caution.


# Function: calculate_burden

calculate_burden <- function(caselist, exposure, grouping_vars = c()) {
  grouping_vars <- grouping_vars[grouping_vars != ""]
  valid_grouping_vars <- intersect(grouping_vars, names(caselist))
  
  df <- caselist %>%
    group_by(across(all_of(c(valid_grouping_vars, "player_id")))) %>%
    summarise(
      n_cases = n(),
      timeloss = sum(timeloss),
      .groups = "drop"
    ) %>%
    mutate(burden = timeloss / exposure * 1000)
  
  group_combinations <- if (length(valid_grouping_vars) > 0) {
    df %>% select(all_of(valid_grouping_vars)) %>% distinct()
  } else {
    tibble(.dummy = 1)
  }
  
  bootSum <- function(data, indices) sum(data[indices], na.rm = TRUE)
  results_list <- list()
  
  for (i in 1:nrow(group_combinations)) {
    current_group <- if (length(valid_grouping_vars) > 0) group_combinations[i, ] else tibble()
    subset_data <- if (length(valid_grouping_vars) > 0) {
      df %>% inner_join(current_group, by = valid_grouping_vars)
    } else {
      df
    }
    
    if (nrow(subset_data) == 0) {
      cat("No data for", if (length(valid_grouping_vars) > 0) paste(current_group, collapse = ", ") else "overall analysis", "\n")
      next
    }
    
    burden_boot <- boot(subset_data$burden, bootSum, R = 100000)
    
    if (all(burden_boot$t == burden_boot$t[1])) {
      cat("No variability in bootstrap distribution for", if (length(valid_grouping_vars) > 0) paste(current_group, collapse = ", ") else "overall analysis", "\n")
      ci_lower <- ci_upper <- NA
    } else {
      ci <- boot.ci(burden_boot, type = "perc", conf = 0.95)
      ci_lower <- ci$perc[4]
      ci_upper <- ci$perc[5]
    }
    
    results_row <- if (length(valid_grouping_vars) > 0) {
      current_group %>%
        mutate(
          burden_rate = sum(subset_data$burden, na.rm = TRUE),
          lower_bound = ci_lower,
          upper_bound = ci_upper
        )
    } else {
      tibble(
        burden_rate = sum(subset_data$burden, na.rm = TRUE),
        lower_bound = ci_lower,
        upper_bound = ci_upper
      )
    }
    
    results_list[[i]] <- results_row
  }
  
  bind_rows(results_list)
}

# Helper: Run and label burden analysis
run_analysis <- function(caselist_input, exposure_input, outcome, definition) {
  calculate_burden(caselist_input, exposure_input) %>%
    mutate(outcome = outcome, definition = definition) %>%
    select(outcome, definition, everything())
}

### Perform analyses ------

# All injuries
all_ma <- run_analysis(caselist, sum(data_exposure$total), "All injuries", "Medical attention")
all_tl <- run_analysis(filter(caselist, timeloss_cat == "timeloss"), sum(data_exposure$total), "All injuries", "Time loss")

# Match injuries
match_ma <- run_analysis(filter(caselist, when_occurred == "Match"), sum(data_exposure$match), "Match injuries", "Medical attention")
match_tl <- run_analysis(filter(caselist, when_occurred == "Match", timeloss_cat == "timeloss"), sum(data_exposure$match), "Match injuries", "Time loss")

# Training injuries
tr_ma <- run_analysis(filter(caselist, when_occurred == "Training"), sum(data_exposure$training), "Training injuries", "Medical attention")
tr_tl <- run_analysis(filter(caselist, when_occurred == "Training", timeloss_cat == "timeloss"), sum(data_exposure$training), "Training injuries", "Time loss")

# Illnesses
all_tl_illness <- run_analysis(
  filter(caselist, timeloss_cat == "timeloss", problem_type == "Illness"),
  player_days,
  "All illnesses",
  "Time loss"
)

# Combine and export results
result_burden <- bind_rows(all_tl, match_tl, tr_tl, all_tl_illness)

result <- left_join(result_incidence, result_burden)

# Save overall results

save_overall_results_excel <- function(result_df, output_path) {
  
  # Define styles
  style_level_1 <- createStyle(textDecoration = "bold", border = "top")
  style_level_2 <- createStyle(indent = 5, halign = "left")
  style_level_3 <- createStyle(textDecoration = "italic", indent = 10)
  style_level_4 <- createStyle(textDecoration = "italic", indent = 15)
  
  # Create workbook and worksheet
  wb <- createWorkbook()
  addWorksheet(wb, "overall_results")
  writeData(wb, "overall_results", result_df)
  
  # Apply conditional formatting based on 'definition' column
  conditionalFormatting(wb, "overall_results", cols = 2:ncol(result_df), rows = 2:1000, rule = '$B2="Medical attention"', style = style_level_2)
  conditionalFormatting(wb, "overall_results", cols = 2:ncol(result_df), rows = 2:1000, rule = '$B2="Time loss"', style = style_level_3)
  
  # Apply bold header
  addStyle(wb, "overall_results", style = style_level_1, rows = 1, cols = 1:ncol(result_df), gridExpand = TRUE)
  
  # Auto-adjust column widths
  setColWidths(wb, "overall_results", cols = 1:ncol(result_df), widths = "auto")
  
  # Save workbook
  saveWorkbook(wb, file = output_path, overwrite = TRUE)
}


save_overall_results_excel(result, make_output_path("overall_results.xlsx"))





# Table 1 ------------

generate_table_1 <- function(caselist_input, exposure_input, output_path) {
  # Incidence and severity
  levels <- list(
    list(level = 1, grouping_vars = c("osiics_15_body_part"), fill = list(osiics_15_tissue = "All", osiics_15_pathology = "All", osiics_15_diagnosis = "All")),
    list(level = 2, grouping_vars = c("osiics_15_body_part", "osiics_15_tissue"), fill = list(osiics_15_pathology = "All", osiics_15_diagnosis = "All")),
    list(level = 3, grouping_vars = c("osiics_15_body_part", "osiics_15_tissue", "osiics_15_pathology"), fill = list(osiics_15_diagnosis = "All")),
    list(level = 4, grouping_vars = c("osiics_15_body_part", "osiics_15_tissue", "osiics_15_pathology", "osiics_15_diagnosis"), fill = list())
  )
  
  incidence_tables <- lapply(levels, function(lvl) {
    calculate_incidence_severity(caselist_input, exposure_input, grouping_vars = lvl$grouping_vars) %>%
      mutate(level = lvl$level) %>%
      mutate(!!!lvl$fill) %>%
      select(level, osiics_15_body_part, osiics_15_tissue, osiics_15_pathology, osiics_15_diagnosis,
             n_cases, incidence_rate, ci_lower, ci_upper, total_timeloss, median_timeloss, q1_timeloss, q3_timeloss)
  })
  
  table_1 <- bind_rows(incidence_tables)
  
  # Burden
  burden_tables <- lapply(levels, function(lvl) {
    calculate_burden(caselist_input, exposure_input, grouping_vars = lvl$grouping_vars) %>%
      mutate(level = lvl$level) %>%
      mutate(!!!lvl$fill) %>%
      select(level, osiics_15_body_part, osiics_15_tissue, osiics_15_pathology, osiics_15_diagnosis,
             burden_rate, lower_bound, upper_bound)
  })
  
  table_burden <- bind_rows(burden_tables)
  
  # Join tables
  table_1 <- left_join(table_1, table_burden)
  
  # Apply factor levels
  table_1$osiics_15_body_part <- factor(table_1$osiics_15_body_part, levels = level_1_order)
  table_1$osiics_15_tissue <- factor(table_1$osiics_15_tissue, levels = level_2_order)
  table_1$osiics_15_pathology <- factor(table_1$osiics_15_pathology, levels = level_3_order)
  table_1$osiics_15_diagnosis <- factor(table_1$osiics_15_diagnosis, levels = level_4_order)
  
  table_1 <- table_1 %>%
    arrange(osiics_15_body_part, osiics_15_tissue, osiics_15_pathology) %>%
    mutate(
      label = case_when(
        level == 1 ~ as.character(osiics_15_body_part),
        level == 2 ~ as.character(osiics_15_tissue),
        level == 3 ~ as.character(osiics_15_pathology),
        level == 4 ~ as.character(osiics_15_diagnosis)
      ),
      sort_order = row_number(),
      area_tissue = paste0(osiics_15_body_part, "_", osiics_15_tissue),
      incidence_ci = paste0(" [", format(round(ci_lower, 2), nsmall = 2), ", ", format(round(ci_upper, 2), nsmall = 2), "]"),
      median_iqr = paste0(" (", q1_timeloss, ", ", q3_timeloss, ")"),
      burden_ci = paste0(" [", format(round(lower_bound, 2), nsmall = 2), ", ", format(round(upper_bound, 2), nsmall = 2), "]")
    ) %>%
    select(level, label, n_cases, incidence_rate, incidence_ci, median_timeloss, median_iqr, burden_rate, burden_ci)
  
  # Save as Excel file
  style_level_1 <- openxlsx::createStyle(textDecoration = "bold", border = "top")
  style_level_2 <- openxlsx::createStyle(indent = 5, halign = "left")
  style_level_3 <- openxlsx::createStyle(textDecoration = "italic", indent = 10)
  style_level_4 <- openxlsx::createStyle(textDecoration = "italic", indent = 15)
  
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "table_1")
  openxlsx::writeData(wb, "table_1", table_1)
  openxlsx::conditionalFormatting(wb, "table_1", cols = 2:9, rows = 2:1000, rule = "$A2=1", style = style_level_1)
  openxlsx::conditionalFormatting(wb, "table_1", cols = 2, rows = 2:1000, rule = "$A2=2", style = style_level_2)
  openxlsx::conditionalFormatting(wb, "table_1", cols = 2, rows = 2:1000, rule = "$A2=3", style = style_level_3)
  openxlsx::conditionalFormatting(wb, "table_1", cols = 2, rows = 2:1000, rule = "$A2=4", style = style_level_4)
  openxlsx::setColWidths(wb, "table_1", cols = 1:9, widths = "auto")
  openxlsx::saveWorkbook(wb, file = output_path, overwrite = TRUE)
}

# Time loss - All injuries
generate_table_1(
  caselist_input = caselist %>% filter(problem_type == "Injury", timeloss_cat == "timeloss"),
  exposure_input = sum(data_exposure$match),
  output_path = make_output_path("table_1_tl_all.xlsx")
)

# Time loss - Match injuries
generate_table_1(
  caselist_input = caselist %>% filter(problem_type == "Injury", timeloss_cat == "timeloss", when_occurred == "Match"),
  exposure_input = sum(data_exposure$match),
  output_path = make_output_path("table_1_tl_match.xlsx")
)

# Time loss - Training injuries
generate_table_1(
  caselist_input = caselist %>% filter(problem_type == "Injury", timeloss_cat == "timeloss", when_occurred == "Training"),
  exposure_input = sum(data_exposure$training),
  output_path = make_output_path("table_1_tl_training.xlsx")
)


# Medical attention - All injuries
generate_table_1(
  caselist_input = caselist %>% filter(problem_type == "Injury"),
  exposure_input = sum(data_exposure$match),
  output_path = make_output_path("table_1_ma_all.xlsx")
)

# Medical attention  - Match injuries
generate_table_1(
  caselist_input = caselist %>% filter(problem_type == "Injury",  when_occurred == "Match"),
  exposure_input = sum(data_exposure$match),
  output_path = make_output_path("table_1_ma_match.xlsx")
)

# Medical attention  - Training injuries
generate_table_1(
  caselist_input = caselist %>% filter(problem_type == "Injury",  when_occurred == "Training"),
  exposure_input = sum(data_exposure$training),
  output_path = make_output_path("table_1_ma_training.xlsx")
)


# Table 2 -----

generate_table_2 <- function(caselist_input, exposure_input, output_path) {
  # Define levels for Table 2
  levels <- list(
    list(level = 2, grouping_vars = c("osiics_15_tissue"), fill = list(osiics_15_pathology = "All")),
    list(level = 3, grouping_vars = c("osiics_15_tissue", "osiics_15_pathology"), fill = list())
  )
  
  # Incidence and severity
  incidence_tables <- lapply(levels, function(lvl) {
    calculate_incidence_severity(caselist_input, exposure_input, grouping_vars = lvl$grouping_vars) %>%
      mutate(level = lvl$level) %>%
      mutate(!!!lvl$fill) %>%
      select(level, osiics_15_tissue, osiics_15_pathology,
             n_cases, incidence_rate, ci_lower, ci_upper,
             total_timeloss, median_timeloss, q1_timeloss, q3_timeloss)
  })
  
  table_2 <- bind_rows(incidence_tables)
  
  # Burden
  burden_tables <- lapply(levels, function(lvl) {
    calculate_burden(caselist_input, exposure_input, grouping_vars = lvl$grouping_vars) %>%
      mutate(level = lvl$level) %>%
      mutate(!!!lvl$fill) %>%
      select(level, osiics_15_tissue, osiics_15_pathology,
             burden_rate, lower_bound, upper_bound)
  })
  
  table_burden <- bind_rows(burden_tables)
  
  # Join tables
  table_2 <- left_join(table_2, table_burden)
  
  # Apply factor levels
  table_2$osiics_15_tissue <- factor(table_2$osiics_15_tissue, levels = level_2_order)
  table_2$osiics_15_pathology <- factor(table_2$osiics_15_pathology, levels = level_3_order)
  
  table_2 <- table_2 %>%
    arrange(osiics_15_tissue, osiics_15_pathology) %>%
    mutate(
      label = case_when(
        level == 2 ~ as.character(osiics_15_tissue),
        level == 3 ~ as.character(osiics_15_pathology)
      ),
      sort_order = row_number(),
      incidence_ci = paste0(" [", format(round(ci_lower, 2), nsmall = 2), ", ", format(round(ci_upper, 2), nsmall = 2), "]"),
      median_iqr = paste0(" (", q1_timeloss, ", ", q3_timeloss, ")"),
      burden_ci = ifelse(is.na(lower_bound) | is.na(upper_bound), "",
                         paste0(" [", format(round(lower_bound, 2), nsmall = 2), ", ", format(round(upper_bound, 2), nsmall = 2), "]"))
    ) %>%
    select(level, label, n_cases, incidence_rate, incidence_ci,
           median_timeloss, median_iqr, burden_rate, burden_ci)
  
  # Save as Excel file
  style_level_2 <- openxlsx::createStyle(textDecoration = "bold", border = "top")
  style_level_3 <- openxlsx::createStyle(textDecoration = "italic", indent = 5)
  
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "table_2")
  openxlsx::writeData(wb, "table_2", table_2)
  openxlsx::conditionalFormatting(wb, "table_2", cols = 2:9, rows = 2:1000, rule = "$A2=2", style = style_level_2)
  openxlsx::conditionalFormatting(wb, "table_2", cols = 2, rows = 2:1000, rule = "$A2=3", style = style_level_3)
  openxlsx::setColWidths(wb, "table_2", cols = 1:9, widths = "auto")
  openxlsx::saveWorkbook(wb, file = output_path, overwrite = TRUE)
}



### Create tables -----
generate_table_2(
  caselist_input = caselist %>% filter(problem_type == "Injury", timeloss_cat == "timeloss"),
  exposure_input = sum(data_exposure$match),
  output_path = make_output_path("table_2_tl_all.xlsx")
)


# Time loss - Match injuries
generate_table_2(
  caselist_input = caselist %>% filter(problem_type == "Injury", timeloss_cat == "timeloss", when_occurred == "Match"),
  exposure_input = sum(data_exposure$match),
  output_path = make_output_path("table_2_tl_match.xlsx")
)

# Time loss - Training injuries
generate_table_2(
  caselist_input = caselist %>% filter(problem_type == "Injury", timeloss_cat == "timeloss", when_occurred == "Training"),
  exposure_input = sum(data_exposure$training),
  output_path = make_output_path("table_2_tl_training.xlsx")
)


# Medical attention - All injuries
generate_table_2(
  caselist_input = caselist %>% filter(problem_type == "Injury"),
  exposure_input = sum(data_exposure$match),
  output_path = make_output_path("table_2_ma_all.xlsx")
)

# Medical attention  - Match injuries
generate_table_2(
  caselist_input = caselist %>% filter(problem_type == "Injury",  when_occurred == "Match"),
  exposure_input = sum(data_exposure$match),
  output_path = make_output_path("table_2_ma_match.xlsx")
)

# Medical attention  - Training injuries
generate_table_2(
  caselist_input = caselist %>% filter(problem_type == "Injury",  when_occurred == "Training"),
  exposure_input = sum(data_exposure$training),
  output_path = make_output_path("table_2_ma_training.xlsx")
)


