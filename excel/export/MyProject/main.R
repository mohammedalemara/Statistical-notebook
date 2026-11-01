
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
form_path <- file.path(source_dir, "s1.xlsx")
output_path <- file.path(output_dir, "selected_translated.xlsx")

# ============================================================
# 4. رابط تصدير البيانات من KoboToolbox
# ============================================================

data_url <- paste0(
  "https://kf.kobotoolbox.org/api/v2/assets/",
  "akFULLhHuubDWuFd2icLho/export-settings/",
  "esU8F5EM7Xv7735XeGnFKaG/data.csv"
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
arabic_col <- "label::Arabic (ar)"
output_label_col <- if (arabic_col %in% names(survey)) {
  arabic_col
} else {
  english_col
}
choice_label_col <- if (arabic_col %in% names(choices)) {
  arabic_col
} else {
  english_col
}

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
  "gov",
  "dg_edu",
  "edu_dir",
  "book_class",
  "sch_env",
  "sch_name_lbl",
  "sch_stat_no",
  "district",
  "subdistrict",
  "sector_lbl",
  "locality_lbl",
  "landmark",
  "est_year",
  "study_lang/ar",
  "study_lang/en",
  "study_lang/oth",
  "is_new",
  "phone",
  "ownership",
  "building_no",
  "gender",
  "independent"
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

# معالجة أسماء الأعمدة المكررة
names(result) <- make.unique(
  new_names,
  sep = "_"
)

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
