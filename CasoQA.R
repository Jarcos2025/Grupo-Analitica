library(dplyr)
library(tidyverse)
library(ggplot2)
library(readxl)

if (!requireNamespace("moments", quietly = TRUE)) {
  install.packages("moments")
}
library(moments)

weekly_visits_data <- read_excel("Web_Analytics.xls", sheet = "Weekly Visits", skip = 4) %>%
  rename(Week = "Week (2008-2009)") %>%
  mutate(Date = as.Date(Week, origin = "1899-12-30")) %>%
  select(Date, Visits, "Unique Visits", Pageviews, "Pages/Visit", "Avg. Time on Site (secs.)", "Bounce Rate", "% New Visits")

financials_data <- read_excel("Web_Analytics.xls", sheet = "Financials", skip = 4) %>%
  rename(Week = "Week (2008-2009)") %>%
  mutate(Date = as.Date(Week, origin = "1899-12-30")) %>%
  select(Date, Revenue, Profit, "Lbs. Sold", Inquiries)

merged_data <- inner_join(weekly_visits_data, financials_data, by = "Date", suffix = c(".visits", ".financials"))

date_initial_end <- as.Date("2008-08-02")
date_pre_promotion_end <- as.Date("2009-02-14")
date_promotion_end <- as.Date("2009-05-09")
date_post_promotion_end <- as.Date("2009-08-29")

merged_data <- merged_data %>%
  mutate(Period = case_when(
    Date <= date_initial_end ~ "Initial",
    Date <= date_pre_promotion_end ~ "Pre-Promotion",
    Date <= date_promotion_end ~ "Promotion",
    Date <= date_post_promotion_end ~ "Post-Promotion",
    TRUE ~ "Other"
  )) %>%
  mutate(Period = factor(Period, levels = c("Initial", "Pre-Promotion", "Promotion", "Post-Promotion", "Other")))

# 1

plot_unique_visits <- ggplot(weekly_visits_data, aes(x = Date, y = `Unique Visits`)) +
  geom_col(fill = "steelblue", color = "black", width = 5) +
  labs(
    title = "Visitas Únicas al Sitio Web de QA a lo largo del Tiempo",
    x = "Fecha",
    y = "Visitas Únicas"
  ) +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"))

print(plot_unique_visits)

plot_revenue <- ggplot(financials_data, aes(x = Date, y = Revenue)) +
  geom_col(fill = "darkgreen", color = "black", width = 5) +
  labs(
    title = "Ingresos a lo largo del Tiempo",
    x = "Fecha",
    y = "Ingresos"
  ) +
  scale_y_continuous(labels = scales::dollar) +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"))

print(plot_revenue)

plot_profit <- ggplot(financials_data, aes(x = Date, y = Profit)) +
  geom_col(fill = "purple", color = "black", width = 5) +
  labs(
    title = "Ganancias a lo largo del Tiempo",
    x = "Fecha",
    y = "Ganancias"
  ) +
  scale_y_continuous(labels = scales::dollar) +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"))

print(plot_profit)

plot_lbs_sold <- ggplot(financials_data, aes(x = Date, y = `Lbs. Sold`)) +
  geom_col(fill = "orange", color = "black", width = 5) +
  labs(
    title = "Libras Vendidas a lo largo del Tiempo",
    x = "Fecha",
    y = "Libras Vendidas"
  ) +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"))

print(plot_lbs_sold)

# 2
variables_to_analyze <- c("Visits", "Unique Visits", "Revenue", "Profit", "Lbs. Sold")
statistic_names <- c("mean", "median", "std. dev", "minimum", "maximum")

for (p in levels(merged_data$Period)) {
  if (p == "Other") next
  
  cat(paste0("\nVISIT AND FINANCIAL SUMMARY MEASURES—", toupper(p), " PERIOD\n"))
  
  period_data <- merged_data %>% filter(Period == p)
  
  stats_for_period <- data.frame(
    Statistic = statistic_names,
    Visits = c(mean(period_data$Visits, na.rm = TRUE),
               median(period_data$Visits, na.rm = TRUE),
               sd(period_data$Visits, na.rm = TRUE),
               min(period_data$Visits, na.rm = TRUE),
               max(period_data$Visits, na.rm = TRUE)),
    `Unique Visits` = c(mean(period_data$`Unique Visits`, na.rm = TRUE),
                        median(period_data$`Unique Visits`, na.rm = TRUE),
                        sd(period_data$`Unique Visits`, na.rm = TRUE),
                        min(period_data$`Unique Visits`, na.rm = TRUE),
                        max(period_data$`Unique Visits`, na.rm = TRUE)),
    Revenue = c(mean(period_data$Revenue, na.rm = TRUE),
                median(period_data$Revenue, na.rm = TRUE),
                sd(period_data$Revenue, na.rm = TRUE),
                min(period_data$Revenue, na.rm = TRUE),
                max(period_data$Revenue, na.rm = TRUE)),
    Profit = c(mean(period_data$Profit, na.rm = TRUE),
               median(period_data$Profit, na.rm = TRUE),
               sd(period_data$Profit, na.rm = TRUE),
               min(period_data$Profit, na.rm = TRUE),
               max(period_data$Profit, na.rm = TRUE)),
    `Lbs. Sold` = c(mean(period_data$`Lbs. Sold`, na.rm = TRUE),
                    median(period_data$`Lbs. Sold`, na.rm = TRUE),
                    sd(period_data$`Lbs. Sold`, na.rm = TRUE),
                    min(period_data$`Lbs. Sold`, na.rm = TRUE),
                    max(period_data$`Lbs. Sold`, na.rm = TRUE))
  )
  print(stats_for_period)
  cat("\n")
}

# 3

# Calcular los promedios por periodo
means_by_period <- merged_data %>%
  group_by(Period) %>%
  summarise(
    Visits = mean(Visits, na.rm = TRUE),
    `Unique Visits` = mean(`Unique Visits`, na.rm = TRUE),
    Revenue = mean(Revenue, na.rm = TRUE),
    Profit = mean(Profit, na.rm = TRUE),
    `Lbs. Sold` = mean(`Lbs. Sold`, na.rm = TRUE)
  ) %>%
  filter(Period != "Other") %>%
  mutate(Period = factor(Period, levels = c("Initial", "Pre-Promotion", "Promotion", "Post-Promotion"))) %>%
  arrange(Period)

cat("TABLA DE MEDIAS POR PERIODO\n")
print(means_by_period)
cat("\n")

summary_long <- merged_data %>%
  group_by(Period) %>%
  summarise(
    Mean_Visits = mean(Visits, na.rm = TRUE),
    Mean_Unique_Visits = mean(`Unique Visits`, na.rm = TRUE),
    Mean_Revenue = mean(Revenue, na.rm = TRUE),
    Mean_Profit = mean(Profit, na.rm = TRUE),
    Mean_Lbs_Sold = mean(`Lbs. Sold`, na.rm = TRUE)
  ) %>%
  pivot_longer(
    cols = starts_with("Mean_"),
    names_to = "Metric",
    values_to = "Mean_Value"
  ) %>%
  mutate(Metric = str_replace(Metric, "Mean_", "")) %>%
  filter(Metric %in% c("Visits", "Unique_Visits", "Revenue", "Profit", "Lbs_Sold"))

plot_mean_visits <- ggplot(filter(summary_long, Metric == "Visits"), aes(x = Period, y = Mean_Value, fill = Period)) +
  geom_col(color = "black") +
  labs(
    title = "Media de Visitas por Periodo",
    x = "Periodo",
    y = "Media de Visitas"
  ) +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"))

print(plot_mean_visits)

plot_mean_unique_visits <- ggplot(filter(summary_long, Metric == "Unique_Visits"), aes(x = Period, y = Mean_Value, fill = Period)) +
  geom_col(color = "black") +
  labs(
    title = "Media de Visitas Únicas por Periodo",
    x = "Periodo",
    y = "Media de Visitas Únicas"
  ) +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"))

print(plot_mean_unique_visits)

# Gráfico de la media de ingresos por periodo
plot_mean_revenue <- ggplot(filter(summary_long, Metric == "Revenue"), aes(x = Period, y = Mean_Value, fill = Period)) +
  geom_col(color = "black") +
  labs(
    title = "Media de Ingresos por Periodo",
    x = "Periodo",
    y = "Media de Ingresos"
  ) +
  scale_y_continuous(labels = scales::dollar) +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"))

print(plot_mean_revenue)

plot_mean_profit <- ggplot(filter(summary_long, Metric == "Profit"), aes(x = Period, y = Mean_Value, fill = Period)) +
  geom_col(color = "black") +
  labs(
    title = "Media de Ganancias por Periodo",
    x = "Periodo",
    y = "Media de Ganancias"
  ) +
  scale_y_continuous(labels = scales::dollar) +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"))

print(plot_mean_profit)

plot_mean_lbs_sold <- ggplot(filter(summary_long, Metric == "Lbs_Sold"), aes(x = Period, y = Mean_Value, fill = Period)) +
  geom_col(color = "black") +
  labs(
    title = "Media de Libras Vendidas por Periodo",
    x = "Periodo",
    y = "Media de Libras Vendidas"
  ) +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"))

print(plot_mean_lbs_sold)

# 5

plot_revenue_lbs_sold_scatter <- ggplot(merged_data, aes(x = `Lbs. Sold`, y = Revenue)) +
  geom_point(color = "darkred", alpha = 0.7) +
  geom_smooth(method = "lm", col = "blue", se = FALSE) +
  labs(
    title = "Diagrama de Dispersión: Ingresos vs. Libras Vendidas",
    x = "Libras Vendidas",
    y = "Ingresos"
  ) +
  scale_y_continuous(labels = scales::dollar) +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"))

print(plot_revenue_lbs_sold_scatter)

correlation_revenue_lbs_sold <- cor(merged_data$Revenue, merged_data$`Lbs. Sold`, use = "complete.obs")
cat(paste0("Coeficiente de Correlación (Revenue vs. Lbs. Sold): ", round(correlation_revenue_lbs_sold, 4), "\n"))

# 6

plot_revenue_visits_scatter <- ggplot(merged_data, aes(x = Visits, y = Revenue)) +
  geom_point(color = "darkblue", alpha = 0.7) +
  geom_smooth(method = "lm", col = "red", se = FALSE) +
  labs(
    title = "Diagrama de Dispersión: Ingresos vs. Visitas",
    x = "Visitas",
    y = "Ingresos"
  ) +
  scale_y_continuous(labels = scales::dollar) +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"))

print(plot_revenue_visits_scatter)

correlation_revenue_visits <- cor(merged_data$Revenue, merged_data$Visits, use = "complete.obs")
cat(paste0("Coeficiente de Correlación (Revenue vs. Visits): ", round(correlation_revenue_visits, 4), "\n"))

# 8

lbs_sold_full_data <- read_excel("Web_Analytics.xls", sheet = "Lbs. Sold", skip = 4) %>%
  rename(Week = `Week`)

lbs_sold_values <- as.numeric(lbs_sold_full_data$`Lbs. Sold`)
lbs_sold_values <- na.omit(lbs_sold_values)

cat("\n## a. Valores de resumen para Lbs. Sold:\n")
mean_lbs_sold <- mean(lbs_sold_values, na.rm = TRUE)
median_lbs_sold <- median(lbs_sold_values, na.rm = TRUE)
sd_lbs_sold <- sd(lbs_sold_values, na.rm = TRUE)
min_lbs_sold <- min(lbs_sold_values, na.rm = TRUE)
max_lbs_sold <- max(lbs_sold_values, na.rm = TRUE)

summary_lbs_sold <- data.frame(
  Statistic = c("mean", "median", "std. dev", "minimum", "maximum"),
  Value = c(mean_lbs_sold, median_lbs_sold, sd_lbs_sold, min_lbs_sold, max_lbs_sold)
)
print(summary_lbs_sold)

cat("\n## b. Histograma de Libras Vendidas:\n")
plot_lbs_sold_histogram <- ggplot(data.frame(Lbs_Sold = lbs_sold_values), aes(x = Lbs_Sold)) +
  geom_histogram(binwidth = 2000, fill = "lightblue", color = "black", alpha = 0.7) +
  labs(
    title = "Histograma de Libras de Material Vendido por Semana",
    x = "Libras Vendidas",
    y = "Frecuencia"
  ) +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"))

print(plot_lbs_sold_histogram)


cat("\n## d. Análisis de la Regla Empírica para Lbs. Sold:\n")

lbs_sold_full_data_with_z <- lbs_sold_full_data %>%
  filter(!is.na(`Lbs. Sold`)) %>%
  mutate(Z_Score = (`Lbs. Sold` - mean_lbs_sold) / sd_lbs_sold)

total_observations <- nrow(lbs_sold_full_data_with_z)

lower_1sd <- mean_lbs_sold - sd_lbs_sold
upper_1sd <- mean_lbs_sold + sd_lbs_sold
actual_1sd <- lbs_sold_full_data_with_z %>%
  filter(`Lbs. Sold` >= lower_1sd & `Lbs. Sold` <= upper_1sd) %>%
  nrow()
theoretical_1sd <- round(total_observations * 0.68)

lower_2sd <- mean_lbs_sold - 2 * sd_lbs_sold
upper_2sd <- mean_lbs_sold + 2 * sd_lbs_sold
actual_2sd <- lbs_sold_full_data_with_z %>%
  filter(`Lbs. Sold` >= lower_2sd & `Lbs. Sold` <= upper_2sd) %>%
  nrow()
theoretical_2sd <- round(total_observations * 0.95)

lower_3sd <- mean_lbs_sold - 3 * sd_lbs_sold
upper_3sd <- mean_lbs_sold + 3 * sd_lbs_sold
actual_3sd <- lbs_sold_full_data_with_z %>%
  filter(`Lbs. Sold` >= lower_3sd & `Lbs. Sold` <= upper_3sd) %>%
  nrow()
theoretical_3sd <- round(total_observations * 0.99)

empirical_rule_table_d <- data.frame(
  Interval = c("mean ± 1 std. dev.", "mean ± 2 std. dev.", "mean ± 3 std. dev."),
  `Theoretical % of Data` = c("68%", "95%", "99%"),
  `Theoretical No. Obs.` = c(theoretical_1sd, theoretical_2sd, theoretical_3sd),
  `Actual No. Obs.` = c(actual_1sd, actual_2sd, actual_3sd)
)
print(empirical_rule_table_d)

cat("\n## e. Análisis refinado de la Regla Empírica para Lbs. Sold:\n")

actual_mean_plus_1sd <- lbs_sold_full_data_with_z %>%
  filter(Z_Score > 0 & Z_Score <= 1) %>%
  nrow()

actual_mean_minus_1sd <- lbs_sold_full_data_with_z %>%
  filter(Z_Score < 0 & Z_Score >= -1) %>%
  nrow()

actual_1sd_to_2sd <- lbs_sold_full_data_with_z %>%
  filter(Z_Score > 1 & Z_Score <= 2) %>%
  nrow()

actual_minus_1sd_to_minus_2sd <- lbs_sold_full_data_with_z %>%
  filter(Z_Score < -1 & Z_Score >= -2) %>%
  nrow()

actual_2sd_to_3sd <- lbs_sold_full_data_with_z %>%
  filter(Z_Score > 2 & Z_Score <= 3) %>%
  nrow()

actual_minus_2sd_to_minus_3sd <- lbs_sold_full_data_with_z %>%
  filter(Z_Score < -2 & Z_Score >= -3) %>%
  nrow()

theoretical_pct_0_to_1sd <- 0.34
theoretical_pct_1sd_to_2sd <- 0.135
theoretical_pct_2sd_to_3sd <- 0.02

empirical_rule_table_e <- data.frame(
  Interval = c("mean + 1 std. dev.", "mean - 1 std. dev.",
               "1 std. dev. to 2 std. dev.", "-1 std. dev. to -2 std. dev.",
               "2 std. dev. to 3 std. dev.", "-2 std. dev. to -3 std. dev."),
  `Theoretical % of Data` = c(paste0(theoretical_pct_0_to_1sd*100, "%"), paste0(theoretical_pct_0_to_1sd*100, "%"),
                              paste0(theoretical_pct_1sd_to_2sd*100, "%"), paste0(theoretical_pct_1sd_to_2sd*100, "%"),
                              paste0(theoretical_pct_2sd_to_3sd*100, "%"), paste0(theoretical_pct_2sd_to_3sd*100, "%")),
  `Theoretical No. Obs.` = c(round(total_observations * theoretical_pct_0_to_1sd), round(total_observations * theoretical_pct_0_to_1sd),
                             round(total_observations * theoretical_pct_1sd_to_2sd), round(total_observations * theoretical_pct_1sd_to_2sd),
                             round(total_observations * theoretical_pct_2sd_to_3sd), round(total_observations * theoretical_pct_2sd_to_3sd)),
  `Actual No. Obs.` = c(actual_mean_plus_1sd, actual_mean_minus_1sd,
                        actual_1sd_to_2sd, actual_minus_1sd_to_minus_2sd,
                        actual_2sd_to_3sd, actual_minus_2sd_to_minus_3sd)
)
print(empirical_rule_table_e)

cat("\n## g. Asimetría (Skewness) y Curtosis para Lbs. Sold:\n")
skewness_lbs_sold <- skewness(lbs_sold_values, na.rm = TRUE)
kurtosis_lbs_sold <- kurtosis(lbs_sold_values, na.rm = TRUE)

cat(paste0("Asimetría (Skewness) para Lbs. Sold: ", round(skewness_lbs_sold, 4), "\n"))
cat(paste0("Curtosis (Kurtosis) para Lbs. Sold: ", round(kurtosis_lbs_sold, 4), "\n"))

# 10

demographics_raw <- read_excel("Web_Analytics.xls", sheet = "Demographics", col_names = FALSE)

extract_and_clean_demographics <- function(data_raw, start_row, end_row, col_names = c("Category", "Visits")) {
  df <- data_raw %>%
    slice(start_row:end_row) %>%
    select(2:3) %>% # Seleccionar las columnas B y C del Excel
    setNames(col_names) %>%
    mutate(Visits = as.numeric(Visits))
  
  if (all(is.na(df$Visits))) {
    warning(paste0("Advertencia: No se encontraron datos válidos de 'Visits' en las filas ", start_row, " a ", end_row, ". El dataframe resultante estará vacío después de filtrar NAs."))
    return(data.frame(Category = character(0), Visits = numeric(0))) # Devolver un DF vacío
  }
  
  df %>%
    filter(!is.na(Visits))
}

traffic_sources_data <- extract_and_clean_demographics(demographics_raw, 9, 12) # Rango corregido

plot_traffic_sources <- ggplot(traffic_sources_data, aes(x = reorder(Category, Visits), y = Visits, fill = Category)) +
  geom_col(color = "black") +
  labs(
    title = "Visitas por Fuente de Tráfico",
    x = "Fuente de Tráfico",
    y = "Número de Visitas"
  ) +
  scale_y_continuous(labels = scales::comma) +
  coord_flip() +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"),
        legend.position = "none") 

print(plot_traffic_sources)

referring_sites_data <- extract_and_clean_demographics(demographics_raw, 15, 24) # Rango corregido

plot_referring_sites <- ggplot(referring_sites_data, aes(x = reorder(Category, Visits), y = Visits, fill = Category)) +
  geom_col(color = "black") +
  labs(
    title = "Top 10 Sitios de Referencia",
    x = "Sitio de Referencia",
    y = "Número de Visitas"
  ) +
  scale_y_continuous(labels = scales::comma) +
  coord_flip() +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"),
        legend.position = "none")

print(plot_referring_sites)

search_engines_data <- extract_and_clean_demographics(demographics_raw, 28, 37) # Rango corregido

plot_search_engines <- ggplot(search_engines_data, aes(x = reorder(Category, Visits), y = Visits, fill = Category)) +
  geom_col(color = "black") +
  labs(
    title = "Top 10 Motores de Búsqueda",
    x = "Motor de Búsqueda",
    y = "Número de Visitas"
  ) +
  scale_y_continuous(labels = scales::comma) +
  coord_flip() +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"),
        legend.position = "none")

print(plot_search_engines)

visits_by_city_data <- extract_and_clean_demographics(demographics_raw, 41, 50) # Rango corregido

plot_visits_by_city <- ggplot(visits_by_city_data, aes(x = reorder(Category, Visits), y = Visits, fill = Category)) +
  geom_col(color = "black") +
  labs(
    title = "Top 10 Regiones Geograficas por Visitas",
    x = "Region",
    y = "Número de Visitas"
  ) +
  scale_y_continuous(labels = scales::comma) +
  coord_flip() +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"),
        legend.position = "none")

print(plot_visits_by_city)

visits_by_country_data <- extract_and_clean_demographics(demographics_raw, 54, 63) # Rango corregido

plot_visits_by_country <- ggplot(visits_by_country_data, aes(x = reorder(Category, Visits), y = Visits, fill = Category)) +
  geom_col(color = "black") +
  labs(
    title = "Top 10 Navegadores por Visitas",
    x = "Navegador",
    y = "Número de Visitas"
  ) +
  scale_y_continuous(labels = scales::comma) +
  coord_flip() +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"),
        legend.position = "none")

print(plot_visits_by_country)


visits_by_browser_data <- extract_and_clean_demographics(demographics_raw, 67, 76) # Rango corregido

plot_visits_by_browser <- ggplot(visits_by_browser_data, aes(x = reorder(Category, Visits), y = Visits, fill = Category)) +
  geom_col(color = "black") +
  labs(
    title = "Top 10 Sistemas Operativos por Visitas",
    x = "Sistema Operativo",
    y = "Número de Visitas"
  ) +
  scale_y_continuous(labels = scales::comma) +
  coord_flip() +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold"),
        legend.position = "none")

print(plot_visits_by_browser)
