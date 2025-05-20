#PAKETTEN

install.packages(c("dplyr", "readxl", "lme4", "lmerTest", "psych"))
library(dplyr)
library(readxl)
library(lme4)
library(lmerTest)
library(psych)

data <- read_excel(file.choose())

names(data)
str(data)

#OMKEREN NEGATIEVE ITEMS
data$Xgeenonnodigewerkpauzes_inv <- 6 - data$Xgeenonnodigewerkpauzes
data$Xveeltijdtelefoon_inv <- 6 - data$Xveeltijdtelefoon

#NIEUWE SAMENGESTELDE VARIABELEN

data$actieve_vermoeidheid <- rowMeans(data[, c("XEmotioneeluitgeput", "Xminderenergie")], na.rm = TRUE)
data$passieve_vermoeidheid <- rowMeans(data[, c("Xafdwalen", "Xconcentreren")], na.rm = TRUE)
data$taakprestatie <- rowMeans(data[, c("Xfunctiebeschrijving", "Xbelangrijkeaspectenverw", "Xformeleeisen")], na.rm = TRUE)
data$ocb <- rowMeans(data[, c("Xhelpencollega", "Xluisterencollega", "Xnuttigeinfocollega",
                              "Xtijdigaangeven", "Xgeenonnodigewerkpauzes_inv", "Xveeltijdtelefoon_inv")], na.rm = TRUE)

#GEMIDDELDEN
daggemiddelden <- data %>%
  group_by(UniqueID, XDatum) %>%
  summarise(
    actieve_vermoeidheid_dag = mean(actieve_vermoeidheid, na.rm = TRUE),
    passieve_vermoeidheid_dag = mean(passieve_vermoeidheid, na.rm = TRUE),
    taakprestatie_dag = mean(taakprestatie, na.rm = TRUE),
    ocb_dag = mean(ocb, na.rm = TRUE),
    aantal_deelnemers_dag = mean(XDeelnemers, na.rm = TRUE),
    inspraak_dag = mean(Xinspraak, na.rm = TRUE),
    verantwoordelijkheid_dag = mean(Xverantwoordelijkheid, na.rm = TRUE),
    .groups = "drop"
  )

#VERGADERINGEN TELLEN

meeting_telling <- data %>%
  group_by(UniqueID, XDatum) %>%
  summarise(
    aantal_online = sum(XVergadervorm == 0, na.rm = TRUE),
    aantal_fysiek = sum(XVergadervorm == 1, na.rm = TRUE),
    .groups = "drop"
  )

#COMBINEREN
dagdata <- left_join(daggemiddelden, meeting_telling, by = c("UniqueID", "XDatum"))



View(dagdata)
summary(dagdata$aantal_online)
summary(dagdata$aantal_fysiek)

table(dagdata$aantal_online, useNA = "ifany")
table(dagdata$aantal_fysiek, useNA = "ifany")

#BETROUWBAARHEID
#ACTIEVE
install.packages("lavaan")
install.packages("semTools")
install.packages("officer")
install.packages("flextable")
install.packages("magrittr")

library(lavaan)
library(semTools)
library(flextable)
library(officer)
library(magrittr)


model_actief <- '
Actief =~ XEmotioneeluitgeput + Xminderenergie
'

fit_actief <- sem(
  model_actief,
  data = data,
  cluster = "UniqueID",
  fixed.x = FALSE,
  estimator = "MLR",
  missing = "fiml",
  likelihood = "normal"
)


reliability_actief <- reliability(fit_actief)
print(reliability_actief)  

alpha_val <- round(as.numeric(reliability_actief["alpha", 1]), 2)
omega_tot <- round(as.numeric(reliability_actief["omega", 1]), 2)
omega_between <- round(as.numeric(reliability_actief["omega2", 1]), 2)
omega_within <- round(as.numeric(reliability_actief["omega3", 1]), 2)

omega_actief_df <- data.frame(
  Coefficient = c("Alpha", "Omega Total", "Omega Between", "Omega Within"),
  Value = c(alpha_val, omega_tot, omega_between, omega_within)
)

doc <- read_docx() %>%
  body_add_par("Tabel: Multilevel Omega voor Actieve Vermoeidheid", style = "heading 1") %>%
  body_add_flextable(flextable(omega_actief_df))

print(doc, target = "Multilevel_Omega_ActieveVermoeidheid.docx")

#Passieve
#Model selectie vragen
model_passief <- '
Passief =~ Xafdwalen + Xconcentreren'

#Berekening
fit_passief <- sem(
  model_passief,
  data = data,
  cluster = "UniqueID",
  fixed.x = FALSE,
  estimator = "MLR",
  missing = "fiml",
  likelihood = "normal"
) 
reliability_passief <- reliability(fit_passief)
print(reliability_passief)   

#Tabel
alpha_val <- round(as.numeric(reliability_passief["alpha", 1]), 2)
omega_tot <- round(as.numeric(reliability_passief["omega", 1]), 2)
omega_between <- round(as.numeric(reliability_passief["omega2", 1]), 2)
omega_within <- round(as.numeric(reliability_passief["omega3", 1]), 2)

omega_passief_df <- data.frame(
  Coefficient = c("Alpha", "Omega Total", "Omega Between", "Omega Within"),
  Value = c(alpha_val, omega_tot, omega_between, omega_within)
)

#Word
doc <- read_docx() %>%
  body_add_par("Tabel: Multilevel Omega voor Passieve Vermoeidheid", style = "heading 1") %>%
  body_add_flextable(flextable(omega_passief_df))

print(doc, target = "Multilevel_Omega_PassieveVermoeidheid.docx")

#Taakprestaties
#Model selectie vragen
model_taak <- '
Taak =~ Xfunctiebeschrijving + Xbelangrijkeaspectenverw + Xformeleeisen
'
fit_taak <- sem(
  model_taak,
  data = data,
  cluster = "UniqueID",
  fixed.x = FALSE,
  estimator = "MLR",
  missing = "fiml",
  likelihood = "normal"
)

#Berekening
reliability_taak <- reliability(fit_taak)
print(reliability_taak) 

#Tabel
alpha_val <- round(as.numeric(reliability_taak["alpha", 1]), 2)
omega_tot <- round(as.numeric(reliability_taak["omega", 1]), 2)
omega_between <- round(as.numeric(reliability_taak["omega2", 1]), 2)
omega_within <- round(as.numeric(reliability_taak["omega3", 1]), 2)

omega_taak_df <- data.frame(
  Coefficient = c("Alpha", "Omega Total", "Omega Between", "Omega Within"),
  Value = c(alpha_val, omega_tot, omega_between, omega_within)
)

#Word
doc <- read_docx() %>%
  body_add_par("Tabel: Multilevel Omega voor Taakprestaties", style = "heading 1") %>%
  body_add_flextable(flextable(omega_taak_df))

print(doc, target = "Multilevel_Omega_Taakprestaties.docx")


#OCB
#Model selectie vragen
model_ocb <- '
OCB =~ Xhelpencollega + Xluisterencollega + Xnuttigeinfocollega +
        Xtijdigaangeven + Xgeenonnodigewerkpauzes_inv + Xveeltijdtelefoon_inv
'
fit_ocb <- sem(
  model_ocb,
  data = data,
  cluster = "UniqueID",
  fixed.x = FALSE,
  estimator = "MLR",
  missing = "fiml",
  likelihood = "normal"
)

#Berekening
reliability_ocb <- reliability(fit_ocb)
print(reliability_ocb)  

#Tabel
alpha_val <- round(as.numeric(reliability_ocb["alpha", 1]), 2)
omega_tot <- round(as.numeric(reliability_ocb["omega", 1]), 2)
omega_between <- round(as.numeric(reliability_ocb["omega2", 1]), 2)
omega_within <- round(as.numeric(reliability_ocb["omega3", 1]), 2)

omega_ocb_df <- data.frame(
  Coefficient = c("Alpha", "Omega Total", "Omega Between", "Omega Within"),
  Value = c(alpha_val, omega_tot, omega_between, omega_within)
)

#Word
doc <- read_docx() %>%
  body_add_par("Tabel: Multilevel Omega voor OCB (volledige schaal)", style = "heading 1") %>%
  body_add_flextable(flextable(omega_ocb_df))

print(doc, target = "Multilevel_Omega_OCB_volledig.docx")


#TESTING WELKE VERLAGEN BETROUWBAARHEID OCB
#Paketten
library(dplyr)
library(psych)

#Selectie items/vragen
items_ocb <- data %>%
  select(Xhelpencollega, Xluisterencollega, Xnuttigeinfocollega, Xtijdigaangeven,
         Xgeenonnodigewerkpauzes_inv, Xveeltijdtelefoon_inv)

#Omega en lading berekenen
omega_resultaat <- omega(items_ocb, nfactors = 1)  

#Tabel aanmaken
ladingen_df <- as.data.frame(omega_resultaat$schmid$sl)  
ladingen_df$Item <- rownames(ladingen_df)  
ladingen_df <- ladingen_df[, c("Item", "g")]  
colnames(ladingen_df) <- c("Item", "Lading op Factor")
ladingen_df$`Lading op Factor` <- round(ladingen_df$`Lading op Factor`, 2)
ft_ladingen <- flextable(ladingen_df)
ft_ladingen <- autofit(ft_ladingen)

#Word
doc <- read_docx() %>%
  body_add_par("Tabel: Factorladingen van OCB-items bij Omega-berekening", style = "heading 1") %>%
  body_add_flextable(ft_ladingen)
print(doc, target = "Omega_OCB_ItemLadingen.docx")

#2 SLECHTSTE WEGLATEN
#Paketten 
library(lavaan)
library(semTools)
library(dplyr)

#Model selectie vragen
model_ocb4 <- '
OCB =~ Xhelpencollega + Xluisterencollega + Xnuttigeinfocollega + Xtijdigaangeven'
fit_ocb4 <- sem(
  model_ocb4,
  data = data,
  cluster = "UniqueID",
  fixed.x = FALSE,
  estimator = "MLR",
  missing = "fiml",
  likelihood = "normal"
)

#Berekening
reliability_ocb4 <- reliability(fit_ocb4)
print(reliability_ocb4)

#Tabel
alpha_val <- round(as.numeric(reliability_ocb4["alpha", 1]), 2)
omega_tot <- round(as.numeric(reliability_ocb4["omega", 1]), 2)
omega_between <- round(as.numeric(reliability_ocb4["omega2", 1]), 2)
omega_within <- round(as.numeric(reliability_ocb4["omega3", 1]), 2)

omega_ocb4_df <- data.frame(
  Coefficient = c("Alpha", "Omega Total", "Omega Between", "Omega Within"),
  Value = c(alpha_val, omega_tot, omega_between, omega_within)
)
#Word
doc <- read_docx()
doc <- body_add_par(doc, "Tabel: Multilevel Omega voor OCB (beste 4 items)", style = "heading 1")
doc <- body_add_flextable(doc, flextable(omega_ocb4_df))
print(doc, target = "Multilevel_Omega_OCB_top4.docx")


#3 BESTE ITEMS

install.packages(c("lavaan", "semTools", "officer", "flextable"))
library(lavaan)
library(semTools)
library(officer)
library(flextable)
library(magrittr)

model_ocb3 <- '
OCB =~ Xhelpencollega + Xluisterencollega + Xnuttigeinfocollega
'

fit_ocb3 <- sem(
  model_ocb3,
  data = data,
  cluster = "UniqueID",
  fixed.x = FALSE,
  estimator = "MLR",
  missing = "fiml",
  likelihood = "normal"
)

reliability_ocb3 <- reliability(fit_ocb3)
print(reliability_ocb3)


alpha_val <- round(as.numeric(reliability_ocb3["alpha", 1]), 2)
omega_tot <- round(as.numeric(reliability_ocb3["omega", 1]), 2)
omega_between <- round(as.numeric(reliability_ocb3["omega2", 1]), 2)
omega_within <- round

omega_ocb3_df <- data.frame(
  Coefficient = c("Alpha", "Omega Total", "Omega Between", "Omega Within"),
  Value = c(alpha_val, omega_tot, omega_between, omega_within)
)

doc <- read_docx() %>%
  body_add_par("Tabel: Multilevel Omega voor OCB (beste 3 items)", style = "heading 1") %>%
  body_add_flextable(flextable(omega_ocb3_df))

print(doc, target = "Multilevel_Omega_OCB_top3.docx")

#AGGREGAAT MET OCB TOP 3
library(dplyr)

ocb_top3_dag <- data %>%
  rowwise() %>%
  mutate(ocb_top3 = mean(c_across(c(Xhelpencollega, Xluisterencollega, Xnuttigeinfocollega)), na.rm = TRUE)) %>%
  ungroup() %>%
  group_by(UniqueID, XDatum) %>%
  summarise(ocb_top3 = mean(ocb_top3, na.rm = TRUE), .groups = "drop")

dagdata <- left_join(dagdata, ocb_top3_dag, by = c("UniqueID", "XDatum"))

#BESCHRIJVENDE STATISTIEK

#Packages
install.packages("psych")
library(psych)
library(apaTables)

#Variabelen
dagdata_selected <- dagdata %>% select(
  actieve_vermoeidheid_dag,
  passieve_vermoeidheid_dag,
  taakprestatie_dag,
  ocb_top3,
  aantal_deelnemers_dag,
  inspraak_dag,
  verantwoordelijkheid_dag
)

#Beschrijvende statistieken en correlaties
beschrijving <- describe(dagdata_selected)
beschrijving_df <- as.data.frame(beschrijving)
correlaties <- cor(dagdata_selected, use = "pairwise.complete.obs")
correlaties_df <- as.data.frame(correlaties)


#Word  
apa.cor.table(
  dagdata_selected,
  filename = "APA_Correlaties_OCBtop3.doc",
  table.number = 1
)

#NULMODELLEN
#Actieve
#Paketten
install.packages("lme4")
library(lme4)


nulmodel_actieve_vermoeidheid <- lmer(actieve_vermoeidheid_dag ~ 1 + (1 | UniqueID), data = dagdata)
summary(nulmodel_actieve_vermoeidheid)

VarCorr(nulmodel_actieve_vermoeidheid)

#ICC
vc <- as.data.frame(VarCorr(nulmodel_actieve_vermoeidheid))
icc <- vc$vcov[1] / (vc$vcov[1] + vc$vcov[2])
icc

#Tonen in tabel 
install.packages("xml2")
install.packages("sjPlot")

library(xml2)
library(sjPlot)

tab_model(nulmodel_actieve_vermoeidheid)
tab_model(nulmodel_actieve_vermoeidheid, file = "nulmodel_actieve_vermoeidheid.doc")

#PAssieve
#Paketten
install.packages("lme4")
library(lme4)

nulmodel_passieve_vermoeidheid <- lmer(passieve_vermoeidheid_dag ~ 1 + (1 | UniqueID), data = dagdata)
summary(nulmodel_passieve_vermoeidheid)

VarCorr(nulmodel_passieve_vermoeidheid)

#ICC
vc <- as.data.frame(VarCorr(nulmodel_passieve_vermoeidheid))
icc <- vc$vcov[1] / (vc$vcov[1] + vc$vcov[2])
icc

#Tonen in tabel 
install.packages("xml2")
install.packages("sjPlot")

library(xml2)
library(sjPlot)

tab_model(nulmodel_passieve_vermoeidheid)
tab_model(nulmodel_passieve_vermoeidheid, file = "nulmodel_passieve_vermoeidheid.doc")

#Taakprestaties
install.packages("lme4")
library(lme4)

install.packages("sjPlot")   
library(sjPlot)

nulmodel_taakprestatie <- lmer(taakprestatie_dag ~ 1 + (1 | UniqueID), data = dagdata)

tab_model(nulmodel_taakprestatie)

tab_model(nulmodel_taakprestatie, file = "nulmodel_taakprestatie.doc")

#OCB
install.packages("lme4")
library(lme4)

install.packages("sjPlot")   
library(sjPlot)

nulmodel_ocb_top3 <- lmer(ocb_top3 ~ 1 + (1 | UniqueID), data = dagdata)

tab_model(nulmodel_ocb_top3)

tab_model(nulmodel_ocb_top3, file = "nulmodel_ocb_top3.doc")

#Tabel van nulmodellen
nulmodel_actieve <- lmer(actieve_vermoeidheid_dag ~ 1 + (1 | UniqueID), data = dagdata)
nulmodel_passieve <- lmer(passieve_vermoeidheid_dag ~ 1 + (1 | UniqueID), data = dagdata)
nulmodel_taakprestatie <- lmer(taakprestatie_dag ~ 1 + (1 | UniqueID), data = dagdata)
nulmodel_ocb_top3 <- lmer(ocb_top3 ~ 1 + (1 | UniqueID), data = dagdata)

#Word
tab_model(
  nulmodel_actieve,
  nulmodel_passieve,
  nulmodel_taakprestatie,
  nulmodel_ocb_top3,
  show.re.var = TRUE,
  show.icc = TRUE,
  show.r2 = TRUE,
  show.se = FALSE,
  digits = 2,
  dv.labels = c("Actieve Vermoeidheid", "Passieve Vermoeidheid", "Taakprestatie", "OCB (Top 3)"),
  title = "Tabel \nParameters nulmodellen voor de afhankelijke variabelen",
  file = "Nulmodellen_OCBtop3.doc"
)

#HYPOTHESE 1 en 2
install.packages("lme4")
install.packages("lmerTest")
install.packages("sjPlot")

library(lme4)
library(lmerTest)
library(sjPlot)

model_actief <- lmer(actieve_vermoeidheid_dag ~ aantal_online + aantal_fysiek + (1 | UniqueID), data = dagdata)
model_passief <- lmer(passieve_vermoeidheid_dag ~ aantal_online + aantal_fysiek + (1 | UniqueID), data = dagdata)
model_taak <- lmer(taakprestatie_dag ~ aantal_online + aantal_fysiek + (1 | UniqueID), data = dagdata)
model_ocb <- lmer(ocb_top3 ~ aantal_online + aantal_fysiek + (1 | UniqueID), data = dagdata)

tab_model(
  model_actief,
  model_passief,
  model_taak,
  model_ocb,
  show.re.var = TRUE,
  show.icc = TRUE,
  show.r2 = TRUE,
  show.se = FALSE,
  digits = 2,
  dv.labels = c("Actieve Vermoeidheid", "Passieve Vermoeidheid", "Taakprestatie", "OCB"),
  title = "Tabel 3: effect van het aantal fysieke en online vergaderingen op vermoeidheid en OCB",
  file = "Tabel3_EffectOnlineFysiek.doc"
)



#Dagtata met telling aantal vergaderingen
library(dplyr)

meeting_telling <- data %>%
  group_by(UniqueID, XDatum) %>%
  summarise(
    aantal_online = sum(XVergadervorm == 0, na.rm = TRUE),
    aantal_fysiek = sum(XVergadervorm == 1, na.rm = TRUE),
    aantal_meetings = n(),
    .groups = "drop"
  )

dagdata_met_telling <- left_join(dagdata, meeting_telling, by = c("UniqueID", "XDatum"))

View(dagdata_met_telling)


#HYPOTHESEN 3 EN 4
#Hypothese 3

install.packages("lme4")
install.packages("lmerTest")
install.packages("sjPlot")

library(lme4)
library(lmerTest)
library(sjPlot)

model_H3 <- lmer(
  taakprestatie_dag ~ actieve_vermoeidheid_dag + (1 | UniqueID),
  data = dagdata_met_telling
)

#Hypothese 4

model_H4 <- lmer(
  taakprestatie_dag ~ aantal_meetings + actieve_vermoeidheid_dag + (1 | UniqueID),
  data = dagdata_met_telling
)

#wordversie
tab_model(
  model_H3,
  model_H4,
  dv.labels = c("H3: Actieve vermoeidheid → taakprestatie", 
                "H4: Mediatie met aantal vergaderingen"),
  title = "Tabel: Modellen voor H3 en H4 – Effect op taakprestatie",
  show.re.var = TRUE,
  show.icc = TRUE,
  show.r2 = TRUE,
  show.se = FALSE,
  digits = 2,
  file = "H3_H4_Taakprestatie.doc"
)


#Hypothese 5

install.packages("lme4")
install.packages("lmerTest")
install.packages("sjPlot")

library(lme4)
library(lmerTest)
library(sjPlot)

model_H5 <- lmer(
  ocb_dag ~ passieve_vermoeidheid_dag + (1 | UniqueID),
  data = dagdata_met_telling
)

#Hypothese 6
model_H6 <- lmer(
  ocb_dag ~ aantal_meetings + passieve_vermoeidheid_dag + (1 | UniqueID),
  data = dagdata_met_telling
)

tab_model(
  model_H5,
  model_H6,
  dv.labels = c(
    "H5: Passieve vermoeidheid → OCB",
    "H6: Mediatie met aantal vergaderingen"
  ),
  title = "Tabel: Modellen voor H5 en H6 – Effect op OCB",
  show.re.var = TRUE,
  show.icc = TRUE,
  show.r2 = TRUE,
  show.se = FALSE,
  digits = 2,
  file = "H5_H6_Modelresultaten_OCB.doc"
)

#CONTROLEVARIABELEN

#Toevoegen baseline meting items uit algemeengroen excel
install.packages(c("readxl", "dplyr"))
library(readxl)
library(dplyr)

data <- read_excel(file.choose())

data <- data %>%
  rowwise() %>%
  mutate(ocb_top3 = mean(c_across(c(Xhelpencollega, Xluisterencollega, Xnuttigeinfocollega)), na.rm = TRUE)) %>%
  ungroup()

dagdata <- data %>%
  group_by(UniqueID, XDatum) %>%
  summarise(
    actieve_vermoeidheid_dag = mean(XEmotioneeluitgeput + Xminderenergie, na.rm = TRUE),
    passieve_vermoeidheid_dag = mean(Xafdwalen + Xconcentreren, na.rm = TRUE),
    taakprestatie_dag = mean(c(Xfunctiebeschrijving, Xbelangrijkeaspectenverw, Xformeleeisen), na.rm = TRUE),
    ocb_top3 = mean(ocb_top3, na.rm = TRUE),
    aantal_deelnemers_dag = mean(XDeelnemers, na.rm = TRUE),
    inspraak_dag = mean(Xinspraak, na.rm = TRUE),
    verantwoordelijkheid_dag = mean(Xverantwoordelijkheid, na.rm = TRUE),
    .groups = "drop"
  )

meeting_telling <- data %>%
  group_by(UniqueID, XDatum) %>%
  summarise(
    aantal_online = sum(XVergadervorm == 0, na.rm = TRUE),
    aantal_fysiek = sum(XVergadervorm == 1, na.rm = TRUE),
    aantal_meetings = n(),
    .groups = "drop"
  )

dagdata_met_telling <- left_join(dagdata, meeting_telling, by = c("UniqueID", "XDatum"))

View(dagdata_met_telling)

baseline <- read_excel(file.choose()) 

dagdata_met_telling <- merge(dagdata_met_telling, baseline, by = "UniqueID")

dagdata_met_telling$Geslacht <- as.factor(dagdata_met_telling$Geslacht)
dagdata_met_telling$Leidinggeven <- as.factor(dagdata_met_telling$Leidinggeven)
dagdata_met_telling$Werkregime <- as.factor(dagdata_met_telling$Werkregime)

names(dagdata_met_telling)
View(dagdata_met_telling)

str(dagdata_met_telling)

#CONTROLEVARIABELEN

library(lme4)
library(lmerTest)
library(sjPlot)

dagdata_met_telling$Geslacht <- as.factor(dagdata_met_telling$Geslacht)
dagdata_met_telling$Leidinggeven <- as.factor(dagdata_met_telling$Leidinggeven)
dagdata_met_telling$Werkregime <- as.factor(dagdata_met_telling$Werkregime)

#Actieve
model_actieve <- lmer(
  actieve_vermoeidheid_dag ~ aantal_online + aantal_fysiek +
    inspraak_dag + verantwoordelijkheid_dag +
    Geslacht + Leidinggeven +
    (1 | UniqueID),
  data = dagdata_met_telling
)

#Passieve
model_passieve <- lmer(
  passieve_vermoeidheid_dag ~ aantal_online + aantal_fysiek +
    inspraak_dag + verantwoordelijkheid_dag +
    Geslacht + Leidinggeven +
    (1 | UniqueID),
  data = dagdata_met_telling
)

#Taakprestaties
model_taak <- lmer(
  taakprestatie_dag ~ aantal_online + aantal_fysiek + actieve_vermoeidheid_dag +
    inspraak_dag + verantwoordelijkheid_dag +
    Geslacht + Leidinggeven +
    (1 | UniqueID),
  data = dagdata_met_telling
)

#OCB
model_ocb <- lmer(
  ocb_top3 ~ aantal_online + aantal_fysiek + passieve_vermoeidheid_dag +
    inspraak_dag + verantwoordelijkheid_dag +
    Geslacht + Werkregime + Leidinggeven +
    (1 | UniqueID),
  data = dagdata_met_telling
)

#Wordversie
tab_model(
  model_actieve,
  model_passieve,
  model_taak,
  model_ocb,
  dv.labels = c("Actieve Vermoeidheid", "Passieve Vermoeidheid", "Taakprestatie", "OCB (Top 3)"),
  title = "Tabel: Effect van aantal fysieke en online vergaderingen op uitkomstvariabelen (met controlevariabelen)",
  file = "Vergadervormen_Gesplitst_met_Controles.doc",
  show.re.var = TRUE,
  show.icc = TRUE,
  show.r2 = TRUE,
  show.se = FALSE,
  digits = 2
)

#EXTRA ANALYSES
#eerste vergadervorm van de dag

library(lme4)
library(lmerTest)
library(sjPlot)
library(dplyr)

#aggregeren
dagdata_eerstevergadering <- data %>%
  group_by(UniqueID, XDatum) %>%
  summarise(
    passieve_vermoeidheid_dag = mean(passieve_vermoeidheid, na.rm = TRUE),
    vergadervorm_dag = first(XVergadervorm),  # 0 = Online, 1 = Fysiek
    .groups = "drop"
  )

dagdata_eerstevergadering$vergadervorm_dag <- factor(
  dagdata_eerstevergadering$vergadervorm_dag,
  levels = c(1, 0),  
  labels = c("Fysiek", "Online")
)

model_passief_eerstevergadering <- lmer(
  passieve_vermoeidheid_dag ~ vergadervorm_dag + (1 | UniqueID),
  data = dagdata_eerstevergadering
)

#Word
tab_model(
  model_passief_eerstevergadering,
  dv.labels = "Passieve vermoeidheid dag",
  title = "Tabel: Effect van de vorm van de eerste vergadering op passieve vermoeidheid",
  file = "Tabel_EersteVergadering_PassieveVermoeidheid.doc",
  show.re.var = TRUE,
  show.icc = TRUE,
  show.r2 = TRUE,
  show.se = FALSE,
  digits = 2
)

#Totaal aantal vergaderingen effect
library(dplyr)
library(lme4)
library(lmerTest)
library(sjPlot)

meeting_telling <- data %>%
  group_by(UniqueID, XDatum) %>%
  summarise(
    aantal_meetings = n(),
    .groups = "drop"
  )

dagdata_met_telling <- left_join(dagdata, meeting_telling, by = c("UniqueID", "XDatum"))

model_aantal_actieve <- lmer(actieve_vermoeidheid_dag ~ aantal_meetings + (1 | UniqueID), data = dagdata_met_telling)
model_aantal_passieve <- lmer(passieve_vermoeidheid_dag ~ aantal_meetings + (1 | UniqueID), data = dagdata_met_telling)
model_aantal_taak <- lmer(taakprestatie_dag ~ aantal_meetings + (1 | UniqueID), data = dagdata_met_telling)
model_aantal_ocb <- lmer(ocb_top3 ~ aantal_meetings + (1 | UniqueID), data = dagdata_met_telling)

tab_model(
  model_aantal_actieve,
  model_aantal_passieve,
  model_aantal_taak,
  model_aantal_ocb,
  dv.labels = c("Actieve Vermoeidheid", "Passieve Vermoeidheid", "Taakprestatie", "OCB (Top 3)"),
  title = "Tabel 8: Effect van het totaal aantal vergaderingen op de afhankelijke variabelen",
  show.re.var = TRUE,
  show.icc = TRUE,
  show.r2 = TRUE,
  show.se = FALSE,
  digits = 2,
  file = "Tabel8_TotaalAantalVergaderingen.doc"
)
