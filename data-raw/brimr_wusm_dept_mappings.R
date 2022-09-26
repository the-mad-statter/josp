brimr_wusm_dept_mappings <- dplyr::tribble(
  ~wusm_dept, ~nih_dept_combining_name,
  "Emergency Medicine", "EMERGENCY MEDICINE",
  "Internal Medicine", "INTERNAL MEDICINE/MEDICINE",
  "Neurology", "NEUROLOGY",
  "Neurosurgery", "NEUROSURGERY",
  "OBGYN", "OBSTETRICS & GYNECOLOGY",
  "Ortho Surgery", "ORTHOPEDICS",
  "Otolaryngology", "OTOLARYNGOLOGY",
  "Radiology & Rad Onc", "RADIATION-DIAGNOSTIC/ONCOLOGY",
  "Anesthesiology", "ANESTHESIOLOGY",
  "Biochemistry", "BIOCHEMISTRY",
  "Cell Biology", "ANATOMY/CELL BIOLOGY",
  "Developmental Biology", "BIOLOGY",
  "Genetics & Genome", "GENETICS",
  "Microbiology", "MICROBIOLOGY/IMMUN/VIROLOGY",
  "Neuroscience", "NEUROSCIENCES",
  "Ophthalmology", "OPHTHALMOLOGY",
  "Pathology", "PATHOLOGY",
  "Pediatrics", "PEDIATRICS",
  "Psychiatry", "PSYCHIATRY",
  "Surgery", "SURGERY"
)

usethis::use_data(brimr_wusm_dept_mappings, overwrite = TRUE)
