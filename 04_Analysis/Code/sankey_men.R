# Sankey plot 

library(tidyverse)
library(extrafont)
#font_import()
loadfonts()
library(ggsankey)

library(ggplot2)
library(dplyr)


chart_data <- 
  read_delim("02_Data/deid/caselist_deid.csv", delim = ",") %>% filter(sex == "men", problem_type == "Injury", timeloss >0, when_occurred != "Other" | is.na(when_occurred)) %>% filter(subsequent_cat != "exacerbation") %>%
  mutate(" " = "All injuries") %>%
  select(" ", onset, when_occurred,  contact, player_action) %>% 
  rename(training_match = when_occurred)

chart_data$onset <-
  recode(
    chart_data$onset,
    "Sudden-onset" = "Sudden onset"
    , "Gradual-onset" = "Gradual onset"
  )

chart_data$contact <-
  recode(
    chart_data$contact,
    "Yes, direct contact (to injured body part)" = "Direct contact"
    , "Yes, indirect contact (to other body part)" = "Indirect contact"
    , "No" = "No contact"
  )

chart_data$player_action <- replace_na(chart_data$player_action, "Other")

chart_data$player_action <-
  recode(
    chart_data$player_action,
    "Running (any speed)" = "Running"
    , "Kicking (any type)" = "Kicking"
    , "Other:" = "Other"
  )
chart_data$player_action <- ifelse(chart_data$player_action  == "", "Other", chart_data$player_action )

unique(chart_data$player_action)

chart_data <- chart_data %>%
  mutate(
    training_match = case_when(
      onset == "Gradual onset" ~ NA,
      onset == "Sudden onset" ~ training_match
    ),
    contact = case_when(
      is.na(training_match) ~ NA,
      onset == "Gradual onset" ~ NA,
      onset != "Gradual onset" ~ contact
    ),
    player_action = case_when(
      is.na(training_match) ~ NA,
      onset == "Gradual onset" ~ NA,
      onset != "Gradual onset" ~ player_action
    ),
    contact = case_when(
      contact == "No" ~ "No contact",
      is.na(training_match) ~ NA,
      onset != "No" ~ contact
    )) %>%
  rename(
    Setting = training_match,
    `Contact type` = contact,
    `Player action` = player_action)

#chart_data$onset <- recode(chart_data$onset, "Sudden" = "Sudden onset", "Gradual" = "Gradual onset")
# Step 1
df <- chart_data %>%
  make_long(" ", onset, Setting,  `Contact type`, `Player action`) %>%
  filter(!is.na(node))
df

df$node = factor(df$node, levels =
                   rev(
                     c("All injuries",
                       
                       "Sudden onset",
                       "Gradual onset",
                       
                       
                       "Training",
                       "Match",
                       
                       "No contact",
                       "Direct contact",
                       "Indirect contact",
                       
                       
                       "Other",
                       "Unknown",
                       "Change of direction",
                       "Falling",
                       "Running",
                       "Kicking",
                       "Controlling the ball",
                       "Landing",
                       "Tackle",
                       "Collision",
                       "Heading",
                       "Hit by ball"
                     )
                   ))


df$next_node = factor(df$next_node, levels =
                        rev(
                          c("All injuries",
                            
                            "Sudden onset",
                            "Gradual onset",
                            
                            "Training",
                            "Match",
                            
                            "No contact",
                            "Direct contact",
                            "Indirect contact",
                            
                            "Other",
                            "Unknown",
                            "Change of direction",
                            "Falling",
                            "Running",
                            "Kicking",
                            "Controlling the ball",
                            "Landing",
                            "Tackle",
                            "Collision",
                            "Heading",
                            "Hit by ball"
                            
                          )
                        ))

#FIFA colours
colours <- c(
  "#00B052",
  "#86B519",
  "#FFCC00",
  "#4994CE",
  "#326295",
  "#FFCC00",
  "#F09316",
  "#8a1a67"
)


data_palette <- read_delim("02_data/raw/palette_nodes.csv", delim = ";")
data_palette$node <- recode(data_palette$node, Reetitive = "All injuries", Acute = "All injuries")
data_palette$node <- recode(data_palette$node,  "Sudden" = "Sudden onset", "Gradual" = "Gradual onset")


colours <- as.vector(data_palette$colour2)
names(colours) <- data_palette$node 
palette_sankey_colour <- scale_colour_manual(name = "node", values =  colours)
palette_sankey_fill <- scale_fill_manual(name = "node", values =  colours)





# Step 2
dagg <- df%>%
  dplyr::group_by(node)%>%
  tally()


# Step 3
df2 <- merge(df, dagg, by.x = 'node', by.y = 'node', all.x = TRUE)



pl <- ggplot(df2, aes(x = x
                      , next_x = next_x
                      , node = node
                      , next_node = next_node
                      , fill = factor(node)
                      
                      , label = paste0(node," n=", n)
)
) 
pl <- pl +geom_sankey(flow.alpha = 0.3, show.legend = TRUE, na.rm = TRUE, space = 2)
pl <- pl +geom_sankey_text(size = 2.2, color = "black", hjust = 0, space = 2, position = position_nudge(x = 0.1))

#pl <- pl +  theme_bw()
pl <- pl + scale_x_discrete(position = "top")

pl <- pl + theme(legend.position = "none", 
                 panel.background = element_rect(fill = NA),
                 axis.text.x = element_text(size = 7,colour = "#326295", margin = margin(t = 0.5, r = 0, b = 0, l = 0), hjust = 0.5, family = "Open Sans Medium"),
                 #plot.margin = unit(c(0.2, 0.2, 0.2, 0.2), "cm"),
                 text=element_text(family="Open Sans")
)
pl <- pl +  theme(axis.title = element_blank()
                  , axis.text.y = element_blank()
                  , axis.text.x = element_blank()
                  , axis.ticks = element_blank()  
                  , panel.grid = element_blank())

#pl <- pl + scale_fill_viridis_d(option = "D")
pl <- pl + annotate("text",x = 6.5, y = 1, label = "", size = 5) # fake point on x axis to increase width to the right
pl <- pl + scale_fill_manual(values = colours)

pl


ggsave("04_Analysis/results/men/sankey_men.pdf", pl, width = 16, height = 9, units = "cm")
ggsave("04_Analysis/results/men/sankey_men.png", pl, device = "png",  width = 16, height = 9, units = "cm", dpi = 600)

