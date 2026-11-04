# ============================================================
# 1. تحميل المكتبات
# ============================================================

library(httr)
library(readr)
library(readxl)
library(dplyr)
library(writexl)

# ============================================================
# 2. إعداد رمز KoboToolbox
# ============================================================

TOKEN <- "ebf64ff6f976640feb272fc5d9c27eab5effeddd"

if (!nzchar(trimws(TOKEN))) {
  stop(
    "لم يتم العثور على رمز Kobo API. ",
    "يرجى إعداد متغير البيئة KOBOTOOLBOX_TOKEN."
  )
}

# ============================================================
# 3. إنشاء مجلدات المشروع
# ============================================================

source_dir <- "source"
output_dir <- "output"

dir.create(
  source_dir,
  showWarnings = FALSE,
  recursive = TRUE
)

dir.create(
  output_dir,
  showWarnings = FALSE,
  recursive = TRUE
)

# مسارات الملفات
data_path <- file.path(source_dir, "data.csv")
form_path <- file.path(source_dir, "S1.xlsx")
output_path <- file.path(output_dir, "selected_translated.xlsx")

# ============================================================
# 4. رابط تصدير البيانات من KoboToolbox
# ============================================================

data_url <- paste0(
  "https://kf.kobotoolbox.org/api/v2/assets/",
  "aJocfurrDafzzN82GZtDV4/export-settings/",
  "es5N5sREa9YyQRcoAU6MKNA/data.csv"
)

# ============================================================
# 5. تنزيل البيانات وحفظها في source/data.csv
# ============================================================

cat("جاري الاتصال بـ KoboToolbox...\n")

response <- GET(
  data_url,
  add_headers(
    Authorization = paste("Token", TOKEN)
  ),
  timeout(120)
)

# التحقق من نجاح الاتصال
if (http_error(response)) {
  stop(
    "فشل تحميل البيانات. رمز HTTP: ",
    status_code(response),
    "\nتحقق من رمز API وصلاحيات الوصول ورابط التصدير."
  )
}

# حفظ الاستجابة في ملف مؤقت
temp_data_path <- tempfile(fileext = ".csv")

writeBin(
  content(response, as = "raw"),
  temp_data_path
)

# تحديد فاصل الأعمدة تلقائيًا بين ; و ,
sample_semicolon <- tryCatch(
  read_delim(
    temp_data_path,
    delim = ";",
    trim_ws = TRUE,
    show_col_types = FALSE,
    name_repair = "minimal",
    n_max = 5
  ),
  error = function(e) NULL
)

sample_comma <- tryCatch(
  read_delim(
    temp_data_path,
    delim = ",",
    trim_ws = TRUE,
    show_col_types = FALSE,
    name_repair = "minimal",
    n_max = 5
  ),
  error = function(e) NULL
)

if (
  is.null(sample_semicolon) ||
  ncol(sample_semicolon) <= 1
) {
  if (!is.null(sample_comma) && ncol(sample_comma) > 1) {
    data_delimiter <- ","
  } else {
    data_delimiter <- ";"
  }
} else {
  data_delimiter <- ";"
}

# قراءة عينة للتحقق من وجود أعمدة البيانات المطلوبة
sample_data <- if (data_delimiter == ";") {
  sample_semicolon
} else {
  sample_comma
}

if (
  is.null(sample_data) ||
  !all(c("book_class", "ownership") %in% names(sample_data))
) {
  unlink(temp_data_path)
  stop(
    "الاستجابة لا تحتوي على أعمدة البيانات المتوقعة. ",
    "تحقق من رابط التصدير وفاصل CSV وصلاحيات الوصول."
  )
}

# حفظ الملف بعد التأكد من صحة بنيته
if (!file.copy(temp_data_path, data_path, overwrite = TRUE)) {
  unlink(temp_data_path)
  stop("تعذر حفظ data.csv داخل مجلد source.")
}

unlink(temp_data_path)

cat(
  "تم حفظ البيانات في: ",
  normalizePath(data_path, winslash = "/", mustWork = TRUE),
  "\n",
  sep = ""
)

# ============================================================
# 6. قراءة البيانات مع السماح بتكرار أسماء الأعمدة
# ============================================================

data <- read_delim(
  data_path,
  delim = data_delimiter,
  trim_ws = TRUE,
  show_col_types = FALSE,
  name_repair = "minimal"
)

cat("عدد السجلات المستوردة:", nrow(data), "\n")

# ============================================================
# 7. تطبيق الفلتر باستخدام القيم الأصلية
# ============================================================

data_filtered <- data %>%
  filter(
    book_class == "kg",
    ownership == "gov"
  )

cat("عدد السجلات بعد الفلتر:", nrow(data_filtered), "\n")

# ============================================================
# 8. قراءة ملف XLSForm
# ============================================================

if (!file.exists(form_path)) {
  stop(
    "ملف s1.xlsx غير موجود. يرجى وضعه داخل مجلد source."
  )
}

survey <- read_excel(
  form_path,
  sheet = "survey",
  .name_repair = "minimal"
)

choices <- read_excel(
  form_path,
  sheet = "choices",
  .name_repair = "minimal"
)

# ============================================================
# 9. التحقق من عمود الترجمة الإنجليزية
# ============================================================

english_col <- "label::English (en)"
arabic_col <- "غير_مستخدم"   # لتعطيل العربية في بقية السكربت
output_label_col <- english_col
choice_label_col <- english_col

if (!all(c("name", "type", english_col) %in% names(survey))) {
  stop(
    "تحقق من ورقة survey والأعمدة name وtype و",
    english_col
  )
}

if (!all(c("list_name", "name", english_col) %in% names(choices))) {
  stop(
    "تحقق من ورقة choices والأعمدة list_name وname و",
    english_col
  )
}

# ============================================================
# 10. تحديد المتغيرات المراد تصديرها
# ============================================================

selected_vars <- c(
  "start",
  "end",
  "form1_lbl",
  "gov",
  "dg_edu",
  "edu_dir",
  "book_class",
  "sch_class",
  "sch_branch/sci",
  "sch_branch/art",
  "sch_env",
  "sch_name",
  "sch_name_lbl",
  "sch_stat_no",
  "district",
  "subdistrict",
  "search_sector",
  "sector",
  "sector_lbl",
  "locality",
  "locality_lbl",
  "landmark",
  "est_year",
  "study_lang/ar",
  "study_lang/en",
  "study_lang/oth",
  "study_lang_lbl",
  "is_new",
  "phone",
  "email",
  "ownership",
  "building_no",
  "gender",
  "shift",
  "independent",
  "parent_sch",
  "parent_sch_lbl",
  "parent_stat_no",
  "shared_type",
  "special_cls_btu",
  "youth_cls",
  "accelerated_cls",
  "registered_kg_lbl",
  "reg_kg_age4_boys",
  "reg_kg_age4_girls",
  "reg_kg_age5_boys",
  "reg_kg_age5_girls",
  "reg_kg_total_boys",
  "reg_kg_total_girls",
  "reg_kg_total_boys_calc",
  "reg_kg_total_girls_calc",
  "classes_kg_lbl",
  "cls_kg_age4_cls_count",
  "cls_kg_age5_cls_count",
  "cls_kg_total_cls_count",
  "cls_kg_total_cls_count_calc",
  "dropouts_kg_lbl",
  "drop_kg_age4_boys",
  "drop_kg_age4_girls",
  "drop_kg_age5_boys",
  "drop_kg_age5_girls",
  "drop_kg_total_boys",
  "drop_kg_total_girls",
  "drop_kg_total_boys_calc",
  "drop_kg_total_girls_calc",
  "new_grade1_lbl",
  "new_g1_age5_boys",
  "new_g1_age5_girls",
  "new_g1_age6_boys",
  "new_g1_age6_girls",
  "new_g1_age7_boys",
  "new_g1_age7_girls",
  "new_g1_age8_boys",
  "new_g1_age8_girls",
  "new_g1_age9_boys",
  "new_g1_age9_girls",
  "new_g1_age10_boys",
  "new_g1_age10_girls",
  "new_g1_total_boys",
  "new_g1_total_girls",
  "new_g1_total_boys_calc",
  "new_g1_total_girls_calc",
  "new_grade7_lbl",
  "new_g7_age12_boys",
  "new_g7_age12_girls",
  "new_g7_age13_boys",
  "new_g7_age13_girls",
  "new_g7_age14_boys",
  "new_g7_age14_girls",
  "new_g7_age15_boys",
  "new_g7_age15_girls",
  "new_g7_age16_boys",
  "new_g7_age16_girls",
  "new_g7_total_boys",
  "new_g7_total_girls",
  "new_g7_total_boys_calc",
  "new_g7_total_girls_calc",
  "new_grade10sci_lbl",
  "new_g10sci_age15_boys",
  "new_g10sci_age15_girls",
  "new_g10sci_age16_boys",
  "new_g10sci_age16_girls",
  "new_g10sci_age17_boys",
  "new_g10sci_age17_girls",
  "new_g10sci_age18_boys",
  "new_g10sci_age18_girls",
  "new_g10sci_age19_boys",
  "new_g10sci_age19_girls",
  "new_g10sci_total_boys",
  "new_g10sci_total_girls",
  "new_g10sci_total_boys_calc",
  "new_g10sci_total_girls_calc",
  "new_grade10art_lbl",
  "new_g10art_age15_boys",
  "new_g10art_age15_girls",
  "new_g10art_age16_boys",
  "new_g10art_age16_girls",
  "new_g10art_age17_boys",
  "new_g10art_age17_girls",
  "new_g10art_age18_boys",
  "new_g10art_age18_girls",
  "new_g10art_age19_boys",
  "new_g10art_age19_girls",
  "new_g10art_total_boys",
  "new_g10art_total_girls",
  "new_g10art_total_boys_calc",
  "new_g10art_total_girls_calc",
  "registered_grade1_lbl",
  "reg_g1_age5_boys",
  "reg_g1_age5_girls",
  "reg_g1_age6_boys",
  "reg_g1_age6_girls",
  "reg_g1_age7_boys",
  "reg_g1_age7_girls",
  "reg_g1_age8_boys",
  "reg_g1_age8_girls",
  "reg_g1_age9_boys",
  "reg_g1_age9_girls",
  "reg_g1_age10_boys",
  "reg_g1_age10_girls",
  "reg_g1_total_boys",
  "reg_g1_total_girls",
  "reg_g1_total_boys_calc",
  "reg_g1_total_girls_calc",
  "registered_grade2_lbl",
  "reg_g2_age7_boys",
  "reg_g2_age7_girls",
  "reg_g2_age8_boys",
  "reg_g2_age8_girls",
  "reg_g2_age9_boys",
  "reg_g2_age9_girls",
  "reg_g2_age10_boys",
  "reg_g2_age10_girls",
  "reg_g2_age11_boys",
  "reg_g2_age11_girls",
  "reg_g2_total_boys",
  "reg_g2_total_girls",
  "reg_g2_total_boys_calc",
  "reg_g2_total_girls_calc",
  "registered_grade3_lbl",
  "reg_g3_age8_boys",
  "reg_g3_age8_girls",
  "reg_g3_age9_boys",
  "reg_g3_age9_girls",
  "reg_g3_age10_boys",
  "reg_g3_age10_girls",
  "reg_g3_age11_boys",
  "reg_g3_age11_girls",
  "reg_g3_age12_boys",
  "reg_g3_age12_girls",
  "reg_g3_total_boys",
  "reg_g3_total_girls",
  "reg_g3_total_boys_calc",
  "reg_g3_total_girls_calc",
  "egistered_grade4_lbl",
  "reg_g4_age9_boys",
  "reg_g4_age9_girls",
  "reg_g4_age10_boys",
  "reg_g4_age10_girls",
  "reg_g4_age11_boys",
  "reg_g4_age11_girls",
  "reg_g4_age12_boys",
  "reg_g4_age12_girls",
  "reg_g4_age13_boys",
  "reg_g4_age13_girls",
  "reg_g4_total_boys",
  "reg_g4_total_girls",
  "reg_g4_total_boys_calc",
  "reg_g4_total_girls_calc",
  "registered_grade5_lbl",
  "reg_g5_age10_boys",
  "reg_g5_age10_girls",
  "reg_g5_age11_boys",
  "reg_g5_age11_girls",
  "reg_g5_age12_boys",
  "reg_g5_age12_girls",
  "reg_g5_age13_boys",
  "reg_g5_age13_girls",
  "reg_g5_age14_boys",
  "reg_g5_age14_girls",
  "reg_g5_total_boys",
  "reg_g5_total_girls",
  "reg_g5_total_boys_calc",
  "reg_g5_total_girls_calc",
  "registered_grade6_lbl",
  "reg_g6_age11_boys",
  "reg_g6_age11_girls",
  "reg_g6_age12_boys",
  "reg_g6_age12_girls",
  "reg_g6_age13_boys",
  "reg_g6_age13_girls",
  "reg_g6_age14_boys",
  "reg_g6_age14_girls",
  "reg_g6_age15_boys",
  "reg_g6_age15_girls",
  "reg_g6_total_boys",
  "reg_g6_total_girls",
  "reg_g6_total_boys_calc",
  "reg_g6_total_girls_calc",
  "registered_grade7_lbl",
  "reg_g7_age12_boys",
  "reg_g7_age12_girls",
  "reg_g7_age13_boys",
  "reg_g7_age13_girls",
  "reg_g7_age14_boys",
  "reg_g7_age14_girls",
  "reg_g7_age15_boys",
  "reg_g7_age15_girls",
  "reg_g7_age16_boys",
  "reg_g7_age16_girls",
  "reg_g7_total_boys",
  "reg_g7_total_girls",
  "reg_g7_total_boys_calc",
  "reg_g7_total_girls_calc",
  "registered_grade8_lbl",
  "reg_g8_age13_boys",
  "reg_g8_age13_girls",
  "reg_g8_age14_boys",
  "reg_g8_age14_girls",
  "reg_g8_age15_boys",
  "reg_g8_age15_girls",
  "reg_g8_age16_boys",
  "reg_g8_age16_girls",
  "reg_g8_age17_boys",
  "reg_g8_age17_girls",
  "reg_g8_total_boys",
  "reg_g8_total_girls",
  "reg_g8_total_boys_calc",
  "reg_g8_total_girls_calc",
  "registered_grade9_lbl",
  "reg_g9_age14_boys",
  "reg_g9_age14_girls",
  "reg_g9_age15_boys",
  "reg_g9_age15_girls",
  "reg_g9_age16_boys",
  "reg_g9_age16_girls",
  "reg_g9_age17_boys",
  "reg_g9_age17_girls",
  "reg_g9_age18_boys",
  "reg_g9_age18_girls",
  "reg_g9_total_boys",
  "reg_g9_total_girls",
  "reg_g9_total_boys_calc",
  "reg_g9_total_girls_calc",
  "registered_grade10sci_lbl",
  "reg_g10sci_age15_boys",
  "reg_g10sci_age15_girls",
  "reg_g10sci_age16_boys",
  "reg_g10sci_age16_girls",
  "reg_g10sci_age17_boys",
  "reg_g10sci_age17_girls",
  "reg_g10sci_age18_boys",
  "reg_g10sci_age18_girls",
  "reg_g10sci_age19_boys",
  "reg_g10sci_age19_girls",
  "reg_g10sci_total_boys",
  "reg_g10sci_total_girls",
  "reg_g10sci_total_boys_calc",
  "reg_g10sci_total_girls_calc",
  "registered_grade10art_lbl",
  "reg_g10art_age15_boys",
  "reg_g10art_age15_girls",
  "reg_g10art_age16_boys",
  "reg_g10art_age16_girls",
  "reg_g10art_age17_boys",
  "reg_g10art_age17_girls",
  "reg_g10art_age18_boys",
  "reg_g10art_age18_girls",
  "reg_g10art_age19_boys",
  "reg_g10art_age19_girls",
  "reg_g10art_total_boys",
  "reg_g10art_total_girls",
  "reg_g10art_total_boys_calc",
  "reg_g10art_total_girls_calc",
  "registered_grade11sci_lbl",
  "reg_g11sci_age16_boys",
  "reg_g11sci_age16_girls",
  "reg_g11sci_age17_boys",
  "reg_g11sci_age17_girls",
  "reg_g11sci_age18_boys",
  "reg_g11sci_age18_girls",
  "reg_g11sci_age19_boys",
  "reg_g11sci_age19_girls",
  "reg_g11sci_age20_boys",
  "reg_g11sci_age20_girls",
  "reg_g11sci_total_boys",
  "reg_g11sci_total_girls",
  "reg_g11sci_total_boys_calc",
  "reg_g11sci_total_girls_calc",
  "registered_grade11art_lbl",
  "reg_g11art_age16_boys",
  "reg_g11art_age16_girls",
  "reg_g11art_age17_boys",
  "reg_g11art_age17_girls",
  "reg_g11art_age18_boys",
  "reg_g11art_age18_girls",
  "reg_g11art_age19_boys",
  "reg_g11art_age19_girls",
  "reg_g11art_age20_boys",
  "reg_g11art_age20_girls",
  "reg_g11art_total_boys",
  "reg_g11art_total_girls",
  "reg_g11art_total_boys_calc",
  "reg_g11art_total_girls_calc",
  "registered_grade12sci_lbl",
  "reg_g12sci_age17_boys",
  "reg_g12sci_age17_girls",
  "reg_g12sci_age18_boys",
  "reg_g12sci_age18_girls",
  "reg_g12sci_age19_boys",
  "reg_g12sci_age19_girls",
  "reg_g12sci_age20_boys",
  "reg_g12sci_age20_girls",
  "reg_g12sci_age21_boys",
  "reg_g12sci_age21_girls",
  "reg_g12sci_total_boys",
  "reg_g12sci_total_girls",
  "reg_g12sci_total_boys_calc",
  "reg_g12sci_total_girls_calc",
  "registered_grade12art_lbl",
  "reg_g12art_age17_boys",
  "reg_g12art_age17_girls",
  "reg_g12art_age18_boys",
  "reg_g12art_age18_girls",
  "reg_g12art_age19_boys",
  "reg_g12art_age19_girls",
  "reg_g12art_age20_boys",
  "reg_g12art_age20_girls",
  "reg_g12art_age21_boys",
  "reg_g12art_age21_girls",
  "reg_g12art_total_boys",
  "reg_g12art_total_girls",
  "reg_g12art_total_boys_calc",
  "reg_g12art_total_girls_calc",
  "classes_lbl",
  "cls_stg1_boys",
  "cls_stg1_girls",
  "cls_stg1_mixed",
  "cls_stg2_boys",
  "cls_stg2_girls",
  "cls_stg2_mixed",
  "cls_stg3_boys",
  "cls_stg3_girls",
  "cls_stg3_mixed",
  "cls_stg4_boys",
  "cls_stg4_girls",
  "cls_stg4_mixed",
  "cls_stg5_boys",
  "cls_stg5_girls",
  "cls_stg5_mixed",
  "cls_stg6_boys",
  "cls_stg6_girls",
  "cls_stg6_mixed",
  "cls_stg7_boys",
  "cls_stg7_girls",
  "cls_stg7_mixed",
  "cls_stg8_boys",
  "cls_stg8_girls",
  "cls_stg8_mixed",
  "cls_stg9_boys",
  "cls_stg9_girls",
  "cls_stg9_mixed",
  "cls_stg10sci_boys",
  "cls_stg10sci_girls",
  "cls_stg10sci_mixed",
  "cls_stg10art_boys",
  "cls_stg10art_girls",
  "cls_stg10art_mixed",
  "cls_stg11sci_boys",
  "cls_stg11sci_girls",
  "cls_stg11sci_mixed",
  "cls_stg11art_boys",
  "cls_stg11art_girls",
  "cls_stg11art_mixed",
  "cls_stg12sci_boys",
  "cls_stg12sci_girls",
  "cls_stg12sci_mixed",
  "cls_stg12art_boys",
  "cls_stg12art_girls",
  "cls_stg12art_mixed",
  "cls_total_boys",
  "cls_total_girls",
  "cls_total_mixed",
  "cls_total_boys_calc",
  "cls_total_girls_calc",
  "cls_total_mixed_calc",
  "dropouts_lbl",
  "drop_stg1_boys",
  "drop_stg1_girls",
  "drop_stg2_boys",
  "drop_stg2_girls",
  "drop_stg3_boys",
  "drop_stg3_girls",
  "drop_stg4_boys",
  "drop_stg4_girls",
  "drop_stg5_boys",
  "drop_stg5_girls",
  "drop_stg6_boys",
  "drop_stg6_girls",
  "drop_stg7_boys",
  "drop_stg7_girls",
  "drop_stg8_boys",
  "drop_stg8_girls",
  "drop_stg9_boys",
  "drop_stg9_girls",
  "drop_stg10sci_boys",
  "drop_stg10sci_girls",
  "drop_stg10art_boys",
  "drop_stg10art_girls",
  "drop_stg11sci_boys",
  "drop_stg11sci_girls",
  "drop_stg11art_boys",
  "drop_stg11art_girls",
  "drop_stg12sci_boys",
  "drop_stg12sci_girls",
  "drop_stg12art_boys",
  "drop_stg12art_girls",
  "drop_total_boys",
  "drop_total_girls",
  "drop_total_boys_calc",
  "drop_total_girls_calc",
  "passed_lbl",
  "pass_stg1_boys",
  "pass_stg1_girls",
  "pass_stg2_boys",
  "pass_stg2_girls",
  "pass_stg3_boys",
  "pass_stg3_girls",
  "pass_stg4_boys",
  "pass_stg4_girls",
  "pass_stg5_boys",
  "pass_stg5_girls",
  "pass_stg6_boys",
  "pass_stg6_girls",
  "pass_stg7_boys",
  "pass_stg7_girls",
  "pass_stg8_boys",
  "pass_stg8_girls",
  "pass_stg9_boys",
  "pass_stg9_girls",
  "pass_stg10sci_boys",
  "pass_stg10sci_girls",
  "pass_stg10art_boys",
  "pass_stg10art_girls",
  "pass_stg11sci_boys",
  "pass_stg11sci_girls",
  "pass_stg11art_boys",
  "pass_stg11art_girls",
  "pass_stg12sci_boys",
  "pass_stg12sci_girls",
  "pass_stg12art_boys",
  "pass_stg12art_girls",
  "pass_total_boys",
  "pass_total_girls",
  "pass_total_boys_calc",
  "pass_total_girls_calc",
  "failed_lbl",
  "fail_stg1_boys",
  "fail_stg1_girls",
  "fail_stg2_boys",
  "fail_stg2_girls",
  "fail_stg3_boys",
  "fail_stg3_girls",
  "fail_stg4_boys",
  "fail_stg4_girls",
  "fail_stg5_boys",
  "fail_stg5_girls",
  "fail_stg6_boys",
  "fail_stg6_girls",
  "fail_stg7_boys",
  "fail_stg7_girls",
  "fail_stg8_boys",
  "fail_stg8_girls",
  "fail_stg9_boys",
  "fail_stg9_girls",
  "fail_stg10sci_boys",
  "fail_stg10sci_girls",
  "fail_stg10art_boys",
  "fail_stg10art_girls",
  "fail_stg11sci_boys",
  "fail_stg11sci_girls",
  "fail_stg11art_boys",
  "fail_stg11art_girls",
  "fail_stg12sci_boys",
  "fail_stg12sci_girls",
  "fail_stg12art_boys",
  "fail_stg12art_girls",
  "fail_total_boys",
  "fail_total_girls",
  "fail_total_boys_calc",
  "fail_total_girls_calc",
  "limit_lbl",
  "limit_stg1_boys",
  "limit_stg1_girls",
  "limit_stg2_boys",
  "limit_stg2_girls",
  "limit_stg3_boys",
  "limit_stg3_girls",
  "limit_stg4_boys",
  "limit_stg4_girls",
  "limit_stg5_boys",
  "limit_stg5_girls",
  "limit_stg6_boys",
  "limit_stg6_girls",
  "limit_stg7_boys",
  "limit_stg7_girls",
  "limit_stg8_boys",
  "limit_stg8_girls",
  "limit_stg9_boys",
  "limit_stg9_girls",
  "limit_stg10sci_boys",
  "limit_stg10sci_girls",
  "limit_stg10art_boys",
  "limit_stg10art_girls",
  "limit_stg11sci_boys",
  "limit_stg11sci_girls",
  "limit_stg11art_boys",
  "limit_stg11art_girls",
  "limit_stg12sci_boys",
  "limit_stg12sci_girls",
  "limit_stg12art_boys",
  "limit_stg12art_girls",
  "limit_total_boys",
  "limit_total_girls",
  "limit_total_boys_calc",
  "limit_total_girls_calc",
  "other_lbl",
  "other_stg1_boys",
  "other_stg1_girls",
  "other_stg2_boys",
  "other_stg2_girls",
  "other_stg3_boys",
  "other_stg3_girls",
  "other_stg4_boys",
  "other_stg4_girls",
  "other_stg5_boys",
  "other_stg5_girls",
  "other_stg6_boys",
  "other_stg6_girls",
  "other_stg7_boys",
  "other_stg7_girls",
  "other_stg8_boys",
  "other_stg8_girls",
  "other_stg9_boys",
  "other_stg9_girls",
  "other_stg10sci_boys",
  "other_stg10sci_girls",
  "other_stg10art_boys",
  "other_stg10art_girls",
  "other_stg11sci_boys",
  "other_stg11sci_girls",
  "other_stg11art_boys",
  "other_stg11art_girls",
  "other_stg12sci_boys",
  "other_stg12sci_girls",
  "other_stg12art_boys",
  "other_stg12art_girls",
  "other_total_boys",
  "other_total_girls",
  "other_total_boys_calc",
  "other_total_girls_calc",
  "youth_registered_stage1_lbl",
  "youth_reg_stg1_age10_boys",
  "youth_reg_stg1_age10_girls",
  "youth_reg_stg1_age11_boys",
  "youth_reg_stg1_age11_girls",
  "youth_reg_stg1_age12_boys",
  "youth_reg_stg1_age12_girls",
  "youth_reg_stg1_age13_boys",
  "youth_reg_stg1_age13_girls",
  "youth_reg_stg1_age14_boys",
  "youth_reg_stg1_age14_girls",
  "youth_reg_stg1_age15_boys",
  "youth_reg_stg1_age15_girls",
  "youth_reg_stg1_total_boys",
  "youth_reg_stg1_total_girls",
  "youth_reg_stg1_total_boys_calc",
  "youth_reg_stg1_total_girls_calc",
  "youth_registered_stage2_lbl",
  "youth_reg_stg2_age10_boys",
  "youth_reg_stg2_age10_girls",
  "youth_reg_stg2_age11_boys",
  "youth_reg_stg2_age11_girls",
  "youth_reg_stg2_age12_boys",
  "youth_reg_stg2_age12_girls",
  "youth_reg_stg2_age13_boys",
  "youth_reg_stg2_age13_girls",
  "youth_reg_stg2_age14_boys",
  "youth_reg_stg2_age14_girls",
  "youth_reg_stg2_age15_boys",
  "youth_reg_stg2_age15_girls",
  "youth_reg_stg2_total_boys",
  "youth_reg_stg2_total_girls",
  "youth_reg_stg2_total_boys_calc",
  "youth_reg_stg2_total_girls_calc",
  "youth_registered_stage3_lbl",
  "youth_reg_stg3_age10_boys",
  "youth_reg_stg3_age10_girls",
  "youth_reg_stg3_age11_boys",
  "youth_reg_stg3_age11_girls",
  "youth_reg_stg3_age12_boys",
  "youth_reg_stg3_age12_girls",
  "youth_reg_stg3_age13_boys",
  "youth_reg_stg3_age13_girls",
  "youth_reg_stg3_age14_boys",
  "youth_reg_stg3_age14_girls",
  "youth_reg_stg3_age15_boys",
  "youth_reg_stg3_age15_girls",
  "youth_reg_stg3_total_boys",
  "youth_reg_stg3_total_girls",
  "youth_reg_stg3_total_boys_calc",
  "youth_reg_stg3_total_girls_calc",
  "youth_registered_stage4_lbl",
  "youth_reg_stg4_age10_boys",
  "youth_reg_stg4_age10_girls",
  "youth_reg_stg4_age11_boys",
  "youth_reg_stg4_age11_girls",
  "youth_reg_stg4_age12_boys",
  "youth_reg_stg4_age12_girls",
  "youth_reg_stg4_age13_boys",
  "youth_reg_stg4_age13_girls",
  "youth_reg_stg4_age14_boys",
  "youth_reg_stg4_age14_girls",
  "youth_reg_stg4_age15_boys",
  "youth_reg_stg4_age15_girls",
  "youth_reg_stg4_total_boys",
  "youth_reg_stg4_total_girls",
  "youth_reg_stg4_total_boys_calc",
  "youth_reg_stg4_total_girls_calc",
  "youth_classes_lbl",
  "youth_cls_stg1_boys",
  "youth_cls_stg1_girls",
  "youth_cls_stg1_mixed",
  "youth_cls_stg2_boys",
  "youth_cls_stg2_girls",
  "youth_cls_stg2_mixed",
  "youth_cls_stg3_boys",
  "youth_cls_stg3_girls",
  "youth_cls_stg3_mixed",
  "youth_cls_stg4_boys",
  "youth_cls_stg4_girls",
  "youth_cls_stg4_mixed",
  "youth_cls_total_boys",
  "youth_cls_total_girls",
  "youth_cls_total_mixed",
  "youth_cls_total_boys_calc",
  "youth_cls_total_girls_calc",
  "youth_cls_total_mixed_calc",
  "youth_dropouts_lbl",
  "youth_drop_stg1_boys",
  "youth_drop_stg1_girls",
  "youth_drop_stg2_boys",
  "youth_drop_stg2_girls",
  "youth_drop_stg3_boys",
  "youth_drop_stg3_girls",
  "youth_drop_stg4_boys",
  "youth_drop_stg4_girls",
  "youth_drop_total_boys",
  "youth_drop_total_girls",
  "youth_drop_total_boys_calc",
  "youth_drop_total_girls_calc",
  "youth_passed_lbl",
  "youth_pass_stg1_boys",
  "youth_pass_stg1_girls",
  "youth_pass_stg2_boys",
  "youth_pass_stg2_girls",
  "youth_pass_stg3_boys",
  "youth_pass_stg3_girls",
  "youth_pass_stg4_boys",
  "youth_pass_stg4_girls",
  "youth_pass_total_boys",
  "youth_pass_total_girls",
  "youth_pass_total_boys_calc",
  "youth_pass_total_girls_calc",
  "youth_failed_lbl",
  "youth_fail_stg1_boys",
  "youth_fail_stg1_girls",
  "youth_fail_stg2_boys",
  "youth_fail_stg2_girls",
  "youth_fail_stg3_boys",
  "youth_fail_stg3_girls",
  "youth_fail_stg4_boys",
  "youth_fail_stg4_girls",
  "youth_fail_total_boys",
  "youth_fail_total_girls",
  "youth_fail_total_boys_calc",
  "youth_fail_total_girls_calc",
  "youth_limit_lbl",
  "youth_limit_stg1_boys",
  "youth_limit_stg1_girls",
  "youth_limit_stg2_boys",
  "youth_limit_stg2_girls",
  "youth_limit_stg3_boys",
  "youth_limit_stg3_girls",
  "youth_limit_stg4_boys",
  "youth_limit_stg4_girls",
  "youth_limit_total_boys",
  "youth_limit_total_girls",
  "youth_limit_total_boys_calc",
  "youth_limit_total_girls_calc",
  "youth_other_lbl",
  "youth_other_stg1_boys",
  "youth_other_stg1_girls",
  "youth_other_stg2_boys",
  "youth_other_stg2_girls",
  "youth_other_stg3_boys",
  "youth_other_stg3_girls",
  "youth_other_stg4_boys",
  "youth_other_stg4_girls",
  "youth_other_total_boys",
  "youth_other_total_girls",
  "youth_other_total_boys_calc",
  "youth_other_total_girls_calc",
  "acc_registered_stage1_lbl",
  "acc_reg_stg1_age12_boys",
  "acc_reg_stg1_age12_girls",
  "acc_reg_stg1_age13_boys",
  "acc_reg_stg1_age13_girls",
  "acc_reg_stg1_age14_boys",
  "acc_reg_stg1_age14_girls",
  "acc_reg_stg1_age15_boys",
  "acc_reg_stg1_age15_girls",
  "acc_reg_stg1_age16_boys",
  "acc_reg_stg1_age16_girls",
  "acc_reg_stg1_age17_boys",
  "acc_reg_stg1_age17_girls",
  "acc_reg_stg1_age18_boys",
  "acc_reg_stg1_age18_girls",
  "acc_reg_stg1_total_boys",
  "acc_reg_stg1_total_girls",
  "acc_reg_stg1_total_boys_calc",
  "acc_reg_stg1_total_girls_calc",
  "acc_registered_stage2_lbl",
  "acc_reg_stg2_age12_boys",
  "acc_reg_stg2_age12_girls",
  "acc_reg_stg2_age13_boys",
  "acc_reg_stg2_age13_girls",
  "acc_reg_stg2_age14_boys",
  "acc_reg_stg2_age14_girls",
  "acc_reg_stg2_age15_boys",
  "acc_reg_stg2_age15_girls",
  "acc_reg_stg2_age16_boys",
  "acc_reg_stg2_age16_girls",
  "acc_reg_stg2_age17_boys",
  "acc_reg_stg2_age17_girls",
  "acc_reg_stg2_age18_boys",
  "acc_reg_stg2_age18_girls",
  "acc_reg_stg2_total_boys",
  "acc_reg_stg2_total_girls",
  "acc_reg_stg2_total_boys_calc",
  "acc_reg_stg2_total_girls_calc",
  "acc_registered_stage3_lbl",
  "acc_reg_stg3_age12_boys",
  "acc_reg_stg3_age12_girls",
  "acc_reg_stg3_age13_boys",
  "acc_reg_stg3_age13_girls",
  "acc_reg_stg3_age14_boys",
  "acc_reg_stg3_age14_girls",
  "acc_reg_stg3_age15_boys",
  "acc_reg_stg3_age15_girls",
  "acc_reg_stg3_age16_boys",
  "acc_reg_stg3_age16_girls",
  "acc_reg_stg3_age17_boys",
  "acc_reg_stg3_age17_girls",
  "acc_reg_stg3_age18_boys",
  "acc_reg_stg3_age18_girls",
  "acc_reg_stg3_total_boys",
  "acc_reg_stg3_total_girls",
  "acc_reg_stg3_total_boys_calc",
  "acc_reg_stg3_total_girls_calc",
  "acc_classes_lbl",
  "acc_cls_stg1_boys",
  "acc_cls_stg1_girls",
  "acc_cls_stg1_mixed",
  "acc_cls_stg2_boys",
  "acc_cls_stg2_girls",
  "acc_cls_stg2_mixed",
  "acc_cls_stg3_boys",
  "acc_cls_stg3_girls",
  "acc_cls_stg3_mixed",
  "acc_cls_total_boys",
  "acc_cls_total_girls",
  "acc_cls_total_mixed",
  "acc_cls_total_boys_calc",
  "acc_cls_total_girls_calc",
  "acc_cls_total_mixed_calc",
  "acc_dropouts_lbl",
  "acc_drop_stg1_boys",
  "acc_drop_stg1_girls",
  "acc_drop_stg2_boys",
  "acc_drop_stg2_girls",
  "acc_drop_stg3_boys",
  "acc_drop_stg3_girls",
  "acc_drop_total_boys",
  "acc_drop_total_girls",
  "acc_drop_total_boys_calc",
  "acc_drop_total_girls_calc",
  "acc_passed_lbl",
  "acc_pass_stg1_boys",
  "acc_pass_stg1_girls",
  "acc_pass_stg2_boys",
  "acc_pass_stg2_girls",
  "acc_pass_stg3_boys",
  "acc_pass_stg3_girls",
  "acc_pass_total_boys",
  "acc_pass_total_girls",
  "acc_pass_total_boys_calc",
  "acc_pass_total_girls_calc",
  "acc_failed_lbl",
  "acc_fail_stg1_boys",
  "acc_fail_stg1_girls",
  "acc_fail_stg2_boys",
  "acc_fail_stg2_girls",
  "acc_fail_stg3_boys",
  "acc_fail_stg3_girls",
  "acc_fail_total_boys",
  "acc_fail_total_girls",
  "acc_fail_total_boys_calc",
  "acc_fail_total_girls_calc",
  "acc_limit_lbl",
  "acc_limit_stg1_boys",
  "acc_limit_stg1_girls",
  "acc_limit_stg2_boys",
  "acc_limit_stg2_girls",
  "acc_limit_stg3_boys",
  "acc_limit_stg3_girls",
  "acc_limit_total_boys",
  "acc_limit_total_girls",
  "acc_limit_total_boys_calc",
  "acc_limit_total_girls_calc",
  "acc_other_lbl",
  "acc_other_stg1_boys",
  "acc_other_stg1_girls",
  "acc_other_stg2_boys",
  "acc_other_stg2_girls",
  "acc_other_stg3_boys",
  "acc_other_stg3_girls",
  "acc_other_total_boys",
  "acc_other_total_girls",
  "acc_other_total_boys_calc",
  "acc_other_total_girls_calc",
  "staff_title_lbl",
  "staff_title_mgr_male",
  "staff_title_mgr_female",
  "staff_title_asst_male",
  "staff_title_asst_female",
  "staff_title_tchr_male",
  "staff_title_tchr_female",
  "staff_title_tchr2_male",
  "staff_title_tchr2_female",
  "staff_title_cslr_male",
  "staff_title_cslr_female",
  "staff_title_total_male_calc",
  "staff_title_total_female_calc",
  "staff_title_total_total_calc",
  "staff_title_mgr_total_calc",
  "staff_title_asst_total_calc",
  "staff_title_tchr_total_calc",
  "staff_title_tchr2_total_calc",
  "staff_title_cslr_total_calc",
  "staff_job_lbl",
  "staff_job_admin_male",
  "staff_job_admin_female",
  "staff_job_acct_male",
  "staff_job_acct_female",
  "staff_job_clerk_male",
  "staff_job_clerk_female",
  "staff_job_lab_asst_male",
  "staff_job_lab_asst_female",
  "staff_job_kgsub_male",
  "staff_job_kgsub_female",
  "staff_job_libr_male",
  "staff_job_libr_female",
  "staff_job_serv_male",
  "staff_job_serv_female",
  "staff_job_guard_male",
  "staff_job_guard_female",
  "staff_job_drvr_male",
  "staff_job_drvr_female",
  "staff_job_farm_male",
  "staff_job_farm_female",
  "staff_job_total_male_calc",
  "staff_job_total_female_calc",
  "staff_job_total_total_calc",
  "staff_job_admin_total_calc",
  "staff_job_acct_total_calc",
  "staff_job_clerk_total_calc",
  "staff_job_lab_asst_total_calc",
  "staff_job_kgsub_total_calc",
  "staff_job_libr_total_calc",
  "staff_job_serv_total_calc",
  "staff_job_guard_total_calc",
  "staff_job_drvr_total_calc",
  "staff_job_farm_total_calc",
  "staff_state_lbl",
  "staff_state_active_male",
  "staff_state_active_female",
  "staff_state_scholar_male",
  "staff_state_scholar_female",
  "staff_state_study_male",
  "staff_state_study_female",
  "staff_state_escort_male",
  "staff_state_escort_female",
  "staff_state_unpaid_male",
  "staff_state_unpaid_female",
  "staff_state_sick_male",
  "staff_state_sick_female",
  "staff_state_maternity_male",
  "staff_state_maternity_female",
  "staff_state_seconded_male",
  "staff_state_seconded_female",
  "staff_state_leave45_male",
  "staff_state_leave45_female",
  "staff_state_total_male_calc",
  "staff_state_total_female_calc",
  "staff_state_total_total_calc",
  "staff_state_active_total_calc",
  "staff_state_scholar_total_calc",
  "staff_state_study_total_calc",
  "staff_state_escort_total_calc",
  "staff_state_unpaid_total_calc",
  "staff_state_sick_total_calc",
  "staff_state_maternity_total_calc",
  "staff_state_seconded_total_calc",
  "staff_state_leave45_total_calc",
  "staff_mar_lbl",
  "staff_mar_male_sgl",
  "staff_mar_male_mar",
  "staff_mar_male_wid",
  "staff_mar_male_div",
  "staff_mar_female_sgl",
  "staff_mar_female_mar",
  "staff_mar_female_wid",
  "staff_mar_female_div",
  "staff_mar_male_total_calc",
  "staff_mar_female_total_calc",
  "staff_mar_total_sgl_calc",
  "staff_mar_total_mar_calc",
  "staff_mar_total_wid_calc",
  "staff_mar_total_div_calc",
  "staff_mar_total_total_calc",
  "staff_serv_lbl",
  "staff_serv_le5_male",
  "staff_serv_le5_female",
  "staff_serv_y6_10_male",
  "staff_serv_y6_10_female",
  "staff_serv_y11_15_male",
  "staff_serv_y11_15_female",
  "staff_serv_ge16_male",
  "staff_serv_ge16_female",
  "staff_serv_total_male_calc",
  "staff_serv_total_female_calc",
  "staff_serv_total_total_calc",
  "staff_serv_le5_total_calc",
  "staff_serv_y6_10_total_calc",
  "staff_serv_y11_15_total_calc",
  "staff_serv_ge16_total_calc",
  "staff_age_lbl",
  "staff_age_le30_male",
  "staff_age_le30_female",
  "staff_age_y30_39_male",
  "staff_age_y30_39_female",
  "staff_age_y40_49_male",
  "staff_age_y40_49_female",
  "staff_age_y50_59_male",
  "staff_age_y50_59_female",
  "staff_age_total_male_calc",
  "staff_age_total_female_calc",
  "staff_age_total_total_calc",
  "staff_age_le30_total_calc",
  "staff_age_y30_39_total_calc",
  "staff_age_y40_49_total_calc",
  "staff_age_y50_59_total_calc",
  "staff_train_lbl",
  "staff_train_male_train",
  "staff_train_female_train",
  "staff_train_total_train_calc",
  "staff_lvl_lbl",
  "staff_lvl_lvl1_male",
  "staff_lvl_lvl1_female",
  "staff_lvl_lvl2_male",
  "staff_lvl_lvl2_female",
  "staff_lvl_total_male_calc",
  "staff_lvl_total_female_calc",
  "staff_lvl_total_total_calc",
  "staff_lvl_lvl1_total_calc",
  "staff_lvl_lvl2_total_calc",
  "staff_qual_lbl",
  "staff_qual_male_train",
  "staff_qual_male_dip_edu",
  "staff_qual_male_dip_art",
  "staff_qual_male_dip",
  "staff_qual_male_ba_edu",
  "staff_qual_male_ba",
  "staff_qual_male_hdip",
  "staff_qual_male_msc",
  "staff_qual_male_phd",
  "staff_qual_male_oth",
  "staff_qual_female_train",
  "staff_qual_female_dip_edu",
  "staff_qual_female_dip_art",
  "staff_qual_female_dip",
  "staff_qual_female_ba_edu",
  "staff_qual_female_ba",
  "staff_qual_female_hdip",
  "staff_qual_female_msc",
  "staff_qual_female_phd",
  "staff_qual_female_oth",
  "staff_qual_male_total_calc",
  "staff_qual_female_total_calc",
  "staff_qual_total_total_calc",
  "staff_qual_total_train_calc",
  "staff_qual_total_dip_edu_calc",
  "staff_qual_total_dip_art_calc",
  "staff_qual_total_dip_calc",
  "staff_qual_total_ba_edu_calc",
  "staff_qual_total_ba_calc",
  "staff_qual_total_hdip_calc",
  "staff_qual_total_msc_calc",
  "staff_qual_total_phd_calc",
  "staff_qual_total_oth_calc",
  "staff_qual_kg_lbl",
  "staff_qual_kg_male_train",
  "staff_qual_kg_male_dip_edu",
  "staff_qual_kg_male_dip_art",
  "staff_qual_kg_male_ba_edu",
  "staff_qual_kg_male_ba_kg",
  "staff_qual_kg_male_ba",
  "staff_qual_kg_male_msc",
  "staff_qual_kg_male_phd",
  "staff_qual_kg_male_oth",
  "staff_qual_kg_female_train",
  "staff_qual_kg_female_dip_edu",
  "staff_qual_kg_female_dip_art",
  "staff_qual_kg_female_ba_edu",
  "staff_qual_kg_female_ba_kg",
  "staff_qual_kg_female_ba",
  "staff_qual_kg_female_msc",
  "staff_qual_kg_female_phd",
  "staff_qual_kg_female_oth",
  "staff_qual_kg_male_total_calc",
  "staff_qual_kg_female_total_calc",
  "staff_qual_kg_total_total_calc",
  "staff_qual_kg_total_train_calc",
  "staff_qual_kg_total_dip_edu_calc",
  "staff_qual_kg_total_dip_art_calc",
  "staff_qual_kg_total_ba_edu_calc",
  "staff_qual_kg_total_ba_kg_calc",
  "staff_qual_kg_total_ba_calc",
  "staff_qual_kg_total_msc_calc",
  "staff_qual_kg_total_phd_calc",
  "staff_qual_kg_total_oth_calc",
  "staff_qual_acc_lbl",
  "staff_qual_acc_male_tts",
  "staff_qual_acc_male_sec",
  "staff_qual_acc_male_voc",
  "staff_qual_acc_male_tti",
  "staff_qual_acc_male_uni_edu",
  "staff_qual_acc_male_uni",
  "staff_qual_acc_male_oth",
  "staff_qual_acc_female_tts",
  "staff_qual_acc_female_sec",
  "staff_qual_acc_female_voc",
  "staff_qual_acc_female_tti",
  "staff_qual_acc_female_uni_edu",
  "staff_qual_acc_female_uni",
  "staff_qual_acc_female_oth",
  "staff_qual_acc_male_total_calc",
  "staff_qual_acc_female_total_calc",
  "staff_qual_acc_total_total_calc",
  "staff_qual_acc_total_tts_calc",
  "staff_qual_acc_total_sec_calc",
  "staff_qual_acc_total_voc_calc",
  "staff_qual_acc_total_tti_calc",
  "staff_qual_acc_total_uni_edu_calc",
  "staff_qual_acc_total_uni_calc",
  "staff_qual_acc_total_oth_calc",
  "staff_spec_lbl",
  "staff_spec_male_islam",
  "staff_spec_male_arab",
  "staff_spec_male_eng",
  "staff_spec_male_kurd",
  "staff_spec_male_math",
  "staff_spec_male_bio",
  "staff_spec_male_chem",
  "staff_spec_male_phys",
  "staff_spec_male_soc",
  "staff_spec_male_hist",
  "staff_spec_male_geo",
  "staff_spec_female_islam",
  "staff_spec_female_arab",
  "staff_spec_female_eng",
  "staff_spec_female_kurd",
  "staff_spec_female_math",
  "staff_spec_female_bio",
  "staff_spec_female_chem",
  "staff_spec_female_phys",
  "staff_spec_female_soc",
  "staff_spec_female_hist",
  "staff_spec_female_geo",
  "staff_spec_male_total_calc",
  "staff_spec_female_total_calc",
  "staff_spec_total_islam_calc",
  "staff_spec_total_arab_calc",
  "staff_spec_total_eng_calc",
  "staff_spec_total_kurd_calc",
  "staff_spec_total_math_calc",
  "staff_spec_total_bio_calc",
  "staff_spec_total_chem_calc",
  "staff_spec_total_phys_calc",
  "staff_spec_total_soc_calc",
  "staff_spec_total_hist_calc",
  "staff_spec_total_geo_calc",
  "staff_spec2_lbl",
  "staff_spec2_male_home",
  "staff_spec2_male_sped",
  "staff_spec2_male_psych",
  "staff_spec2_male_art",
  "staff_spec2_male_sport",
  "staff_spec2_male_comp",
  "staff_spec2_male_turk",
  "staff_spec2_male_syr",
  "staff_spec2_male_oth",
  "staff_spec2_male_hist",
  "staff_spec2_female_home",
  "staff_spec2_female_sped",
  "staff_spec2_female_psych",
  "staff_spec2_female_art",
  "staff_spec2_female_sport",
  "staff_spec2_female_comp",
  "staff_spec2_female_turk",
  "staff_spec2_female_syr",
  "staff_spec2_female_oth",
  "staff_spec2_female_hist",
  "staff_spec2_male_total_calc",
  "staff_spec2_female_total_calc",
  "staff_spec2_total_home_calc",
  "staff_spec2_total_sped_calc",
  "staff_spec2_total_psych_calc",
  "staff_spec2_total_art_calc",
  "staff_spec2_total_sport_calc",
  "staff_spec2_total_comp_calc",
  "staff_spec2_total_turk_calc",
  "staff_spec2_total_syr_calc",
  "staff_spec2_total_oth_calc",
  "staff_spec2_total_hist_calc",
  "staff_spec2_total_total_calc",
  "staff_spec_sec_lbl",
  "staff_spec_sec_male_islam",
  "staff_spec_sec_male_arab",
  "staff_spec_sec_male_eng",
  "staff_spec_sec_male_kurd",
  "staff_spec_sec_male_math",
  "staff_spec_sec_male_phys",
  "staff_spec_sec_male_chem",
  "staff_spec_sec_male_bio",
  "staff_spec_sec_male_soc",
  "staff_spec_sec_male_hist",
  "staff_spec_sec_male_geo",
  "staff_spec_sec_female_islam",
  "staff_spec_sec_female_arab",
  "staff_spec_sec_female_eng",
  "staff_spec_sec_female_kurd",
  "staff_spec_sec_female_math",
  "staff_spec_sec_female_phys",
  "staff_spec_sec_female_chem",
  "staff_spec_sec_female_bio",
  "staff_spec_sec_female_soc",
  "staff_spec_sec_female_hist",
  "staff_spec_sec_female_geo",
  "staff_spec_sec_male_total_calc",
  "staff_spec_sec_female_total_calc",
  "staff_spec_sec_total_islam_calc",
  "staff_spec_sec_total_arab_calc",
  "staff_spec_sec_total_eng_calc",
  "staff_spec_sec_total_kurd_calc",
  "staff_spec_sec_total_math_calc",
  "staff_spec_sec_total_phys_calc",
  "staff_spec_sec_total_chem_calc",
  "staff_spec_sec_total_bio_calc",
  "staff_spec_sec_total_soc_calc",
  "staff_spec_sec_total_hist_calc",
  "staff_spec_sec_total_geo_calc",
  "staff_spec_sec2_lbl",
  "staff_spec_sec2_male_mgmt",
  "staff_spec_sec2_male_econ",
  "staff_spec_sec2_male_art",
  "staff_spec_sec2_male_sport",
  "staff_spec_sec2_male_psych",
  "staff_spec_sec2_male_sped",
  "staff_spec_sec2_male_home",
  "staff_spec_sec2_male_comp",
  "staff_spec_sec2_male_turk",
  "staff_spec_sec2_male_syr",
  "staff_spec_sec2_male_hist",
  "staff_spec_sec2_female_mgmt",
  "staff_spec_sec2_female_econ",
  "staff_spec_sec2_female_art",
  "staff_spec_sec2_female_sport",
  "staff_spec_sec2_female_psych",
  "staff_spec_sec2_female_sped",
  "staff_spec_sec2_female_home",
  "staff_spec_sec2_female_comp",
  "staff_spec_sec2_female_turk",
  "staff_spec_sec2_female_syr",
  "staff_spec_sec2_female_hist",
  "staff_spec_sec2_male_total_calc",
  "staff_spec_sec2_female_total_calc",
  "staff_spec_sec2_total_mgmt_calc",
  "staff_spec_sec2_total_econ_calc",
  "staff_spec_sec2_total_art_calc",
  "staff_spec_sec2_total_sport_calc",
  "staff_spec_sec2_total_psych_calc",
  "staff_spec_sec2_total_sped_calc",
  "staff_spec_sec2_total_home_calc",
  "staff_spec_sec2_total_comp_calc",
  "staff_spec_sec2_total_turk_calc",
  "staff_spec_sec2_total_syr_calc",
  "staff_spec_sec2_total_hist_calc",
  "staff_spec_sec2_total_total_calc",
  "app_lbl",
  "app_guid",
  "app_feed",
  "app_net",
  "app_comp",
  "app_for_lang",
  "app_for_lang_name",
  "app_local_lang",
  "app_sbm",
  "app_child",
  "app_early",
  "app_pos_edu",
  "app_reg_non_iq",
  "registered_non_iraqi_lbl",
  "reg_non_irq_kg4_boys",
  "reg_non_irq_kg4_girls",
  "reg_non_irq_kg5_boys",
  "reg_non_irq_kg5_girls",
  "reg_non_irq_stg1_boys",
  "reg_non_irq_stg1_girls",
  "reg_non_irq_stg2_boys",
  "reg_non_irq_stg2_girls",
  "reg_non_irq_stg3_boys",
  "reg_non_irq_stg3_girls",
  "reg_non_irq_stg4_boys",
  "reg_non_irq_stg4_girls",
  "reg_non_irq_stg5_boys",
  "reg_non_irq_stg5_girls",
  "reg_non_irq_stg6_boys",
  "reg_non_irq_stg6_girls",
  "reg_non_irq_stg7_boys",
  "reg_non_irq_stg7_girls",
  "reg_non_irq_stg8_boys",
  "reg_non_irq_stg8_girls",
  "reg_non_irq_stg9_boys",
  "reg_non_irq_stg9_girls",
  "reg_non_irq_stg10sci_boys",
  "reg_non_irq_stg10sci_girls",
  "reg_non_irq_stg10art_boys",
  "reg_non_irq_stg10art_girls",
  "reg_non_irq_stg11sci_boys",
  "reg_non_irq_stg11sci_girls",
  "reg_non_irq_stg11art_boys",
  "reg_non_irq_stg11art_girls",
  "reg_non_irq_stg12sci_boys",
  "reg_non_irq_stg12sci_girls",
  "reg_non_irq_stg12art_boys",
  "reg_non_irq_stg12art_girls",
  "reg_non_irq_total_boys",
  "reg_non_irq_total_girls",
  "reg_non_irq_total_boys_calc",
  "reg_non_irq_total_girls_calc",
  "show_special_cls/c1",
  "show_special_cls/c2",
  "show_special_cls/c3",
  "show_special_cls/c4",
  "show_special_cls/c5",
  "show_special_cls/c6",
  "special_cls_grade1_lbl",
  "special_cls_g1_age5_boys",
  "special_cls_g1_age5_girls",
  "special_cls_g1_age6_boys",
  "special_cls_g1_age6_girls",
  "special_cls_g1_age7_boys",
  "special_cls_g1_age7_girls",
  "special_cls_g1_age8_boys",
  "special_cls_g1_age8_girls",
  "special_cls_g1_age9_boys",
  "special_cls_g1_age9_girls",
  "special_cls_g1_age10_boys",
  "special_cls_g1_age10_girls",
  "special_cls_g1_total_boys",
  "special_cls_g1_total_girls",
  "special_cls_g1_total_boys_calc",
  "special_cls_g1_total_girls_calc",
  "special_cls_grade2_lbl",
  "special_cls_g2_age7_boys",
  "special_cls_g2_age7_girls",
  "special_cls_g2_age8_boys",
  "special_cls_g2_age8_girls",
  "special_cls_g2_age9_boys",
  "special_cls_g2_age9_girls",
  "special_cls_g2_age10_boys",
  "special_cls_g2_age10_girls",
  "special_cls_g2_age11_boys",
  "special_cls_g2_age11_girls",
  "special_cls_g2_total_boys",
  "special_cls_g2_total_girls",
  "special_cls_g2_total_boys_calc",
  "special_cls_g2_total_girls_calc",
  "special_cls_grade3_lbl",
  "special_cls_g3_age8_boys",
  "special_cls_g3_age8_girls",
  "special_cls_g3_age9_boys",
  "special_cls_g3_age9_girls",
  "special_cls_g3_age10_boys",
  "special_cls_g3_age10_girls",
  "special_cls_g3_age11_boys",
  "special_cls_g3_age11_girls",
  "special_cls_g3_age12_boys",
  "special_cls_g3_age12_girls",
  "special_cls_g3_total_boys",
  "special_cls_g3_total_girls",
  "special_cls_g3_total_boys_calc",
  "special_cls_g3_total_girls_calc",
  "special_cls_grade4_lbl",
  "special_cls_g4_age9_boys",
  "special_cls_g4_age9_girls",
  "special_cls_g4_age10_boys",
  "special_cls_g4_age10_girls",
  "special_cls_g4_age11_boys",
  "special_cls_g4_age11_girls",
  "special_cls_g4_age12_boys",
  "special_cls_g4_age12_girls",
  "special_cls_g4_age13_boys",
  "special_cls_g4_age13_girls",
  "special_cls_g4_total_boys",
  "special_cls_g4_total_girls",
  "special_cls_g4_total_boys_calc",
  "special_cls_g4_total_girls_calc",
  "special_cls_grade5_lbl",
  "special_cls_g5_age10_boys",
  "special_cls_g5_age10_girls",
  "special_cls_g5_age11_boys",
  "special_cls_g5_age11_girls",
  "special_cls_g5_age12_boys",
  "special_cls_g5_age12_girls",
  "special_cls_g5_age13_boys",
  "special_cls_g5_age13_girls",
  "special_cls_g5_age14_boys",
  "special_cls_g5_age14_girls",
  "special_cls_g5_total_boys",
  "special_cls_g5_total_girls",
  "special_cls_g5_total_boys_calc",
  "special_cls_g5_total_girls_calc",
  "special_cls_grade6_lbl",
  "special_cls_g6_age11_boys",
  "special_cls_g6_age11_girls",
  "special_cls_g6_age12_boys",
  "special_cls_g6_age12_girls",
  "special_cls_g6_age13_boys",
  "special_cls_g6_age13_girls",
  "special_cls_g6_age14_boys",
  "special_cls_g6_age14_girls",
  "special_cls_g6_age15_boys",
  "special_cls_g6_age15_girls",
  "special_cls_g6_total_boys",
  "special_cls_g6_total_girls",
  "special_cls_g6_total_boys_calc",
  "special_cls_g6_total_girls_calc",
  "special_classes_lbl",
  "special_cls_stg1_boys",
  "special_cls_stg1_girls",
  "special_cls_stg1_mixed",
  "special_cls_stg2_boys",
  "special_cls_stg2_girls",
  "special_cls_stg2_mixed",
  "special_cls_stg3_boys",
  "special_cls_stg3_girls",
  "special_cls_stg3_mixed",
  "special_cls_stg4_boys",
  "special_cls_stg4_girls",
  "special_cls_stg4_mixed",
  "special_cls_stg5_boys",
  "special_cls_stg5_girls",
  "special_cls_stg5_mixed",
  "special_cls_stg6_boys",
  "special_cls_stg6_girls",
  "special_cls_stg6_mixed",
  "special_cls_total_boys",
  "special_cls_total_girls",
  "special_cls_total_mixed",
  "special_cls_total_boys_calc",
  "special_cls_total_girls_calc",
  "special_cls_total_mixed_calc",
  "disability_students_lbl",
  "dis_slow_lr_boys",
  "dis_slow_lr_girls",
  "dis_phys_dis_boys",
  "dis_phys_dis_girls",
  "dis_vis_dis_boys",
  "dis_vis_dis_girls",
  "dis_hear_dis_boys",
  "dis_hear_dis_girls",
  "dis_intel_dis_boys",
  "dis_intel_dis_girls",
  "dis_psych_case_boys",
  "dis_psych_case_girls",
  "dis_total_boys",
  "dis_total_girls",
  "dis_total_boys_calc",
  "dis_total_girls_calc",
  "staff_special_lbl",
  "staff_special_male_islam",
  "staff_special_male_arab",
  "staff_special_male_eng",
  "staff_special_male_kurd",
  "staff_special_male_math",
  "staff_special_male_bio",
  "staff_special_male_chem",
  "staff_special_male_phys",
  "staff_special_male_soc",
  "staff_special_male_hist",
  "staff_special_male_geo",
  "staff_special_female_islam",
  "staff_special_female_arab",
  "staff_special_female_eng",
  "staff_special_female_kurd",
  "staff_special_female_math",
  "staff_special_female_bio",
  "staff_special_female_chem",
  "staff_special_female_phys",
  "staff_special_female_soc",
  "staff_special_female_hist",
  "staff_special_female_geo",
  "staff_special_male_total_calc",
  "staff_special_female_total_calc",
  "staff_special_total_islam_calc",
  "staff_special_total_arab_calc",
  "staff_special_total_eng_calc",
  "staff_special_total_kurd_calc",
  "staff_special_total_math_calc",
  "staff_special_total_bio_calc",
  "staff_special_total_chem_calc",
  "staff_special_total_phys_calc",
  "staff_special_total_soc_calc",
  "staff_special_total_hist_calc",
  "staff_special_total_geot_calc",
  "staff_special2_lbl",
  "staff_special2_male_home",
  "staff_special2_male_sped",
  "staff_special2_male_psych",
  "staff_special2_male_art",
  "staff_special2_male_sport",
  "staff_special2_male_comp",
  "staff_special2_male_turk",
  "staff_special2_male_syr",
  "staff_special2_male_oth",
  "staff_special2_male_hist",
  "staff_special2_female_home",
  "staff_special2_female_sped",
  "staff_special2_female_psych",
  "staff_special2_female_art",
  "staff_special2_female_sport",
  "staff_special2_female_comp",
  "staff_special2_female_turk",
  "staff_special2_female_syr",
  "staff_special2_female_oth",
  "staff_special2_female_hist",
  "staff_special2_male_total_calc",
  "staff_special2_female_total_calc",
  "staff_special2_total_home_calc",
  "staff_special2_total_sped_calc",
  "staff_special2_total_psych_calc",
  "staff_special2_total_art_calc",
  "staff_special2_total_sport_calc",
  "staff_special2_total_comp_calc",
  "staff_special2_total_turk_calc",
  "staff_special2_total_syr_calc",
  "staff_special2_total_oth_calc",
  "staff_special2_total_hist_calc",
  "staff_special2_total_total_calc",
  "school_building_lbl",
  "bld_type",
  "const_year",
  "plot_no",
  "sector_no",
  "exp_space",
  "floors",
  "need_cls",
  "extra_cls",
  "capacity",
  "classrooms",
  "other_rooms",
  "total_rooms",
  "bld_material",
  "has_san",
  "san_cond",
  "san_cap",
  "drink_water",
  "water_src",
  "has_sewer",
  "sewer_type",
  "has_fence",
  "fence_type",
  "fence_cond",
  "sports_field",
  "garden",
  "activity_hall",
  "hall_type",
  "occupancy",
  "bld_owner",
  "bld_cond",
  "eq_furn_lbl",
  "has_eq",
  "dbl_desk_good",
  "dbl_desk_damg",
  "sgl_desk_good",
  "sgl_desk_damg",
  "boards_good",
  "boards_damg",
  "computers_good",
  "computers_damg",
  "printers_good",
  "printers_damg",
  "scanners_good",
  "scanners_damg",
  "library_lbl",
  "has_lib",
  "lib_staff",
  "lib_type",
  "book_cnt",
  "mag_story_cnt",
  "total_items",
  "school_labs_lbl",
  "has_lab",
  "lab_type/sci",
  "lab_type/phy",
  "lab_type/chem",
  "lab_type/bio",
  "lab_type/mix",
  "lab_type/lang",
  "lab_type/comp",
  "sci_lab_ind",
  "sci_lab_room",
  "phy_lab_ind",
  "phy_lab_room",
  "chem_lab_ind",
  "chem_lab_room",
  "bio_lab_ind",
  "bio_lab_room",
  "mix_lab_ind",
  "mix_lab_room",
  "lang_lab_ind",
  "lang_lab_room",
  "comp_lab_ind",
  "comp_lab_room",
  "oth_info_lbl",
  "mgr_name",
  "asst_name",
  "asst_mobile",
  "guest_sch1",
  "guest_sch2",
  "mgr_notes",
  "_submitted_by"
)

# الاحتفاظ بالمتغيرات الموجودة في البيانات فقط
selected_vars <- intersect(
  selected_vars,
  names(data_filtered)
)

if (length(selected_vars) == 0) {
  stop("لم يتم العثور على المتغيرات المحددة في البيانات.")
}

result <- data_filtered %>%
  select(all_of(selected_vars))

cat(
  "عدد الأعمدة المحددة للتصدير:",
  length(selected_vars),
  "\n"
)

# ============================================================
# 11. ترجمة قيم select_one و select_multiple
# ============================================================

for (i in seq_len(nrow(survey))) {
  
  var_name <- as.character(survey$name[i])
  question_type <- as.character(survey$type[i])
  
  if (
    is.na(var_name) ||
    !var_name %in% names(result) ||
    is.na(question_type)
  ) {
    next
  }
  
  # تحديد نوع السؤال
  if (grepl("^select_one\\s+", question_type)) {
    
    list_name <- sub(
      "^select_one\\s+",
      "",
      question_type
    )
    
    multiple <- FALSE
    
  } else if (grepl("^select_multiple\\s+", question_type)) {
    
    list_name <- sub(
      "^select_multiple\\s+",
      "",
      question_type
    )
    
    multiple <- TRUE
    
  } else {
    next
  }
  
  # أخذ اسم القائمة فقط، دون إعدادات الأعمدة الاختيارية
  list_name <- strsplit(
    trimws(list_name),
    "\\s+"
  )[[1]][1]
  
  # استخراج الاختيارات التابعة للقائمة
  ch <- choices[
    !is.na(choices$list_name) &
      as.character(choices$list_name) == list_name,
    ,
    drop = FALSE
  ]
  
  if (nrow(ch) == 0) {
    next
  }
  
  codes <- as.character(ch$name)
  labels <- as.character(ch[[choice_label_col]])
  
  # استخدم الإنجليزية كخيار احتياطي إذا كانت التسمية العربية فارغة.
  bad <- is.na(labels) | !nzchar(trimws(labels))
  if (choice_label_col == arabic_col && english_col %in% names(ch)) {
    english_labels <- as.character(ch[[english_col]])
    use_english <- bad & !is.na(english_labels) & nzchar(trimws(english_labels))
    labels[use_english] <- english_labels[use_english]
  }
  
  # إذا بقيت التسمية فارغة، نحتفظ بالرمز.
  bad <- is.na(labels) | !nzchar(trimws(labels))
  labels[bad] <- codes[bad]
  
  valid <- !is.na(codes) & nzchar(codes)
  
  value_map <- setNames(
    labels[valid],
    codes[valid]
  )
  
  original_values <- as.character(result[[var_name]])
  
  if (!multiple) {
    
    # ترجمة الاختيار الواحد
    idx <- !is.na(original_values) &
      original_values %in% names(value_map)
    
    original_values[idx] <- unname(
      value_map[original_values[idx]]
    )
    
  } else {
    
    # ترجمة الاختيارات المتعددة
    for (j in seq_along(original_values)) {
      
      val <- original_values[j]
      
      if (is.na(val) || !nzchar(trimws(val))) {
        next
      }
      
      codes_j <- strsplit(
        trimws(val),
        "\\s+"
      )[[1]]
      
      labels_j <- codes_j
      matched <- codes_j %in% names(value_map)
      
      labels_j[matched] <- unname(
        value_map[codes_j[matched]]
      )
      
      original_values[j] <- paste(
        labels_j,
        collapse = ", "
      )
    }
  }
  
  result[[var_name]] <- original_values
}

# ============================================================
# 11-b. ترجمة أعمدة الاختيارات المتعددة المنقسمة
# ============================================================
# قد يصدر Kobo كل خيار متعدد كعمود مستقل بالشكل question/choice.
# مثال: study_lang/ar حيث study_lang اسم السؤال و ar رمز الخيار.
for (col_name in names(result)) {
  if (!grepl("/", col_name, fixed = TRUE)) {
    next
  }
  
  parts <- strsplit(col_name, "/", fixed = TRUE)[[1]]
  if (length(parts) != 2L) {
    next
  }
  
  question_name <- parts[1]
  choice_code <- parts[2]
  survey_row <- which(as.character(survey$name) == question_name)
  
  if (length(survey_row) == 0L) {
    next
  }
  
  question_type <- as.character(survey$type[survey_row[1]])
  if (is.na(question_type) ||
      !grepl("^select_multiple\\s+", question_type)) {
    next
  }
  
  list_name <- sub("^select_multiple\\s+", "", question_type)
  list_name <- strsplit(trimws(list_name), "\\s+")[[1]][1]
  
  ch <- choices[
    !is.na(choices$list_name) &
      as.character(choices$list_name) == list_name &
      as.character(choices$name) == choice_code,
    ,
    drop = FALSE
  ]
  
  if (nrow(ch) == 0L) {
    next
  }
  
  # نستخدم تسمية السؤال والخيار العربية، ثم الإنجليزية عند غياب العربية.
  question_label <- if (arabic_col %in% names(survey)) {
    as.character(survey[[arabic_col]][survey_row[1]])
  } else {
    NA_character_
  }
  if (is.na(question_label) || !nzchar(trimws(question_label))) {
    question_label <- as.character(survey[[english_col]][survey_row[1]])
  }
  
  choice_label <- as.character(ch[[choice_label_col]][1])
  if (is.na(choice_label) || !nzchar(trimws(choice_label))) {
    fallback_col <- if (output_label_col == arabic_col) english_col else arabic_col
    choice_label <- if (fallback_col %in% names(choices)) {
      as.character(ch[[fallback_col]][1])
    } else {
      NA_character_
    }
  }
  
  if (!is.na(choice_label) && nzchar(trimws(choice_label))) {
    # نترجم اسم السؤال ورمز الخيار إلى العربية عند توفرها.
    if (is.na(question_label) || !nzchar(trimws(question_label))) {
      question_label <- question_name
    }
    names(result)[names(result) == col_name] <- paste0(
      question_label, "/", choice_label
    )
  }
}

# ============================================================
# 12. ترجمة أسماء الأعمدة إلى الإنجليزية
# ============================================================

original_names <- names(result)

survey_names <- as.character(survey$name)
survey_labels <- as.character(survey[[output_label_col]])

valid_labels <- !is.na(survey_names) &
  nzchar(trimws(survey_names)) &
  !is.na(survey_labels) &
  nzchar(trimws(survey_labels))

label_map <- setNames(
  survey_labels[valid_labels],
  survey_names[valid_labels]
)

new_names <- original_names

for (i in seq_along(original_names)) {
  
  old_name <- original_names[i]
  
  if (old_name %in% names(label_map)) {
    new_names[i] <- label_map[[old_name]]
  }
}

# تطبيق الأسماء الجديدة مع السماح بتكرارها
# نحوّل النتيجة إلى data.frame عادي لأنه يقبل الأسماء المكررة
result <- as.data.frame(result, check.names = FALSE)
names(result) <- new_names

# ============================================================
# 13. تصدير ملف Excel النهائي
# ============================================================

write_xlsx(
  result,
  output_path
)

cat(
  "\nتمت العملية بنجاح.\n",
  "ملف البيانات: ",
  normalizePath(data_path, winslash = "/", mustWork = TRUE),
  "\n",
  "ملف Excel النهائي: ",
  normalizePath(output_path, winslash = "/", mustWork = TRUE),
  "\n",
  "عدد السجلات النهائية: ",
  nrow(result),
  "\n",
  sep = ""
)

# ============================================================
# 14. عرض النتائج
# ============================================================

View(result)