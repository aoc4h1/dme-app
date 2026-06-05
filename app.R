library(shiny)
library(httr)
library(jsonlite)
library(here)
library(dplyr)
library(tidyr)
library(stringi)
library(stringr)
library(rio)
library(purrr)
library(openxlsx)
library(bslib)
library(bsicons)
library(logger)
library(rlang)

# --- Verziószám ---

aktualis_verzioszam <- "v1.0.1."
aktualis_datum <- "2026.06.05."


# --- Segédfüggvények ---

# Függvény generáló
write_dynamic_formula <- function(wb, sheet, formula_template, rows, start_col, start_row) {
  n_placeholders <- stringr::str_count(formula_template, "%d")
  formulas <- sapply(rows, function(r) {
    do.call(sprintf, c(list(fmt = formula_template), as.list(rep(r, n_placeholders))))
  })
  writeFormula(wb, sheet, x = formulas, startCol = start_col, startRow = start_row)
}

# Feltételes formázás (hibák/figyelmeztetések pirosítására)
apply_cond_format <- function(wb, sheet, cols, rows, rule, bg_color = "#FFC7CE", font_color = "black") {
  conditionalFormatting(
    wb, sheet = sheet, cols = cols, rows = rows, rule = rule,
    style = createStyle(fontColour = font_color, bgFill = bg_color)
  )
}

# Legördülő listák hozzáadása
add_dropdown <- function(wb, sheet, cols, rows, value) {
  # Csak akkor adjuk hozzá, ha van érvényes sor
  if (length(rows) > 0 && rows[1] > 0) {
    dataValidation(wb, sheet = sheet, cols = cols, rows = rows, type = "list", value = value)
  }
}

# Hiperhivatkozások ÉS az üres (nincs adat) sorok formázása egyben!
add_mtmt_hyperlinks <- function(wb, sheet, df, id_col, start_row) {
  # Ha van adat és nem az "üres" üzenet van az első sorban
  if (nrow(df) > 0 && !grepl("^A jelölt", df[[1]][1])) {
    linkek <- paste0("https://m2.mtmt.hu/api/publication/", df[[id_col]])
    megjelenitett_szoveg <- gsub('"', "'", df[[1]])
    formulas <- paste0('HYPERLINK("', linkek, '", "', megjelenitett_szoveg, '")')
    
    writeFormula(wb, sheet = sheet, x = formulas, startCol = 1, startRow = start_row)
    
    link_style <- createStyle(fontColour = "#0000FF", textDecoration = "underline")
    addStyle(wb, sheet = sheet, style = link_style, 
             rows = start_row:(start_row + nrow(df) - 1), 
             cols = 1, gridExpand = TRUE)
  } else {
    # Ha nincs adat (csak a hibaüzenet), piros és dőlt betűs stílus az 1. oszlopban
    addStyle(wb, sheet = sheet, 
             style = createStyle(textDecoration = c("bold", "italic"), fontColour = "#FF0000"),
             cols = 1, rows = start_row, gridExpand = FALSE)
  }
}


# --- Loggolás és egyéb ---
# Logolás konfigurálása (fájlba és konzolra)
log_layout(layout_glue_colors)
log_threshold(DEBUG)
# Opcionális: log_appender(appender_file("app_debug.log"))

# Ikonok ellenőrzése
has_icons <- requireNamespace("bsicons", quietly = TRUE)


# --- UI ---

ui <- page_fillable(
  title = "Doktori Minimum Ellenőrző",
  theme = bs_theme(
    version = 5, 
    bootswatch = "flatly", 
    base_font = font_google("Open Sans", wght = c(400, 700)) # 400 a normál, 700 a félkövér
  ),  
  div(
    style = "display: flex; flex-direction: column; justify-content: center; align-items: center; min-height: 90vh; padding: 20px;",
    
    # Fejléc
    div(
      style = "text-align: center; margin-bottom: 30px;",
      h1("Doktori Minimum Ellenőrző", class = "display-4"),
      p("MTMT alapú publikációs és hivatkozási összesítő", class = "mb-1", style = "font-size: 2.2rem;"),
      p("IX. Osztály, Szociológiai Tudományos Bizottság", class = "mb-1", style = "font-size: 2rem;"),
      tags$br(),
      p(
        class = "text-muted small",
        style = "font-size: 0.9rem;",
        "A kalkulációk a ",
        tags$a(
          href = "https://mta.hu/data/dokumentumok/doktori_tanacs/IX.%20Osztaly/2019/9GJO_DoktoriMinimumkovetelmenyekTara_20190628tol.pdf",
          "Doktori Minimumkövetelmények Tárának",
          target = "_blank"
        ),
        " megfelelően készültek el a IX. Osztály Szociológiai Tudományos Bizottságra vonatkozóan.",
        
        tags$br(),
        
        tags$span(
          class = "fw-bold", 
          style = "font-weight: bold !important; color: #2c3e50 !important;",
          "A kalkulátor nem hivatalos program, a kitöltő saját felelőssége, hogy a kapott pontszámokat ellenőrizze. A kalkulációért a program alkotói nem vállalnak felelősséget."
        ),
        
        tags$br(),
        
        tags$span(class = "fw-bold", paste0("Jelenlegi verzió: ", aktualis_verzioszam)), 
        paste0(" (Utolsó frissítés: ", aktualis_datum, ")"),
        
        tags$br(),
        
        "Technikai probléma esetén kérjük, írjon a ",
        tags$a(
          href = "mailto:kmetty.zoltan@tk.hu",
          "kmetty.zoltan@tk.hu"
        ),
        " e-mail címre."
      )
    ),
    
    # Középső kártya az inputtal
    card(
      style = "width: 100%; max-width: 500px; box-shadow: 0 4px 15px rgba(0,0,0,0.1);",
      card_header(
        class = "bg-primary text-white text-center",
        if(has_icons) bs_icon("person-badge") else NULL,
        " Szerző azonosítása"
      ),
      card_body(
        div(
          style = "padding: 10px;",
          numericInput("mtmtid", "MTMT Szerzői ID megadása:", value = "", width = "100%"),
          br(),
          # 1. A gomb, ami elindítja a számolást
          actionButton("generateData", "Adatok feldolgozása", class = "btn-primary btn-lg w-100", icon = icon("gears")),
          br(), br(),
          # 2. Ide kerül dinamikusan a letöltés gomb, miután kész az Excel
          uiOutput("downloadBtnUI")
        )
      ),
      card_footer(
        class = "text-center",
        verbatimTextOutput("status_log", placeholder = TRUE)
      )
    ),
    
    # Használati útmutató
    div(
      style = "width: 100%; max-width: 500px; margin-top: 20px;",
      accordion(
        accordion_panel(
          "Használati útmutató",
          icon = if(has_icons) bs_icon("info-circle") else NULL,
          tags$ol(
            tags$li("Írja be a 8 jegyű MTMT azonosítót."),
            tags$li("Kattintson az 'Adatok feldolgozása' gombra."),
            tags$li("A rendszer pár perc alatt elkészíti az Ön teljesítményét tartalmazó fájlt."),
            tags$li("Kattintson a 'Kész Excel letöltése' gombra.")
          ),
          hr(), # Elválasztó vonal
          
          # ÚJ SZÖVEG MÓDOSÍTOTT STÍLUSSAL
          p(
            class = "text-muted small mb-3 text-start", # text-start a balra igazításért
            style = "font-size: 0.9rem;",
            "Mivel az MTMT rendszere nem tartalmazza a pontszámításhoz szükséges összes adatot, a pontos kalkulációhoz a jelölt kiegészítéseire is szükség van. Az alábbi útmutató segít az elkészült Excel-fájl értelmezésében, és pontosan megmutatja, hol szükséges manuálisan megadni a hiányzó információkat."
          ),
          
          # KATTINTHATÓ GOMB AZ ÚTMUTATÓHOZ (Shiny beépített gomb)
          downloadButton(
            outputId = "downloadGuide", 
            label = " DME - Felhasználói Útmutató.docx letöltése",
            class = "btn btn-outline-info w-100",
            icon = icon("file-word")
          )
        ),
        
        # ÚJ VERZIÓTÖRTÉNET SZEKCIÓ
        accordion_panel(
          "Verziótörténet",
          icon = if(has_icons) bs_icon("clock-history") else NULL,
          div(
            class = "text-start",
            tags$h6(class = "fw-bold", "v1.0.1 (2026.06.05.)"),
            tags$ul(
              class = "small",
              tags$li(tags$strong("Hivatkozások:"), " A 'Norvég lista' oszlop adatkonverziós hibájának javítása a 4_hivatkozasok lapfülön."),
              tags$li(tags$strong("Könyvek besorolása:"), " Pontosított logika az 1_konyv táblázatban. Ha egy könyvnek 2-nél több szerzője van, az automatikusan a '3.1 További pontozott közlemények' kategóriába kerül. Ha 2 vagy kevesebb, az 1_konyv lapfülön is megjelenik.")
            )
            # Ide jöhetnek majd a jövőbeli verziók, pl:
            # hr(),
            # tags$h6(class = "fw-bold", "v1.00 (2026.06.04.)"),
            # tags$ul(class = "small", tags$li("Béta verzió indulása."))
          )
        )
      )
    )
  )
)

# --- Szerver logika (Server) ---
server <- function(input, output, session) {
  
  # Felhasználói útmutató letöltésének kezelése
  output$downloadGuide <- downloadHandler(
    filename = function() {
      # Ezen a néven kapja meg a felhasználó a böngészőből
      "DME_Felhasznaloi_Utmutato.docx" 
    },
    content = function(file) {
      # Ellenőrizzük, hogy az R tényleg látja-e a fájlt
      if (!file.exists("dme_utmutato.docx")) {
        showNotification("Hiba: A fájl nem található a szerveren!", type = "error")
      } else {
        # Ha látja, átmásolja a letöltési útvonalra
        file.copy("dme_utmutato.docx", file)
      }
    }
  )
  
  log_text <- reactiveVal("Készen áll.")
  output$status_log <- renderText({ log_text() })
  
  # 1. Egy reaktív változó, ami a háttérben legenerált Excel fájl elérési útját tárolja
  ready_file_path <- reactiveVal(NULL)
  
  # 2. A számolás elindítása, amikor a felhasználó a "Adatok feldolgozása" gombra kattint
  observeEvent(input$generateData, {
    
    # Ha újra kattintanak, elrejtjük a korábbi letöltés gombot
    ready_file_path(NULL)
    
    # Generálunk egy egyedi, ideiglenes fájlnevet a szerveren
    temp_xlsx <- tempfile(pattern = paste0("mtmt_output_", input$mtmtid, "_"), fileext = ".xlsx")
    
    withProgress(message = 'Adatok feldolgozása...', value = 0, {
      tryCatch({
        session_id <- substr(session$token, 1, 6)
        
        # ---------------------------------------------------------
        # 1. Inputok
        # ---------------------------------------------------------
        
        log_debug("[{session_id}] Könyvtárak és útvonalak ellenőrzése")
        mtmtid <- input$mtmtid
        output_path <- "mtmt_output.xlsx"
        # Eredeti fájlnév a kódodból
        journal_data_path <- "SzocTB_20230601.xlsx"
        
        if(!file.exists(output_path)) {
          log_error("[{session_id}] Hiányzó template: {output_path}")
          stop("A template fájl nem található!")
        }
        
        # ---------------------------------------------------------
        # 2. INFO sheet
        # ---------------------------------------------------------
        
        incProgress(0.1, detail = "Journal adatok...")
        log_info("[{session_id}] Journal adatok beolvasása")
        hazai <- suppressMessages(import(journal_data_path, sheet = "Hazai"))
        hazai$nemzetkozi <- F
        nemzetkozi <- suppressMessages(import(journal_data_path, sheet = "Nemzetközi"))
        nemzetkozi$nemzetkozi <- T
        journaldata <- rbind(hazai, nemzetkozi)
        rm(hazai, nemzetkozi)
        log_info("[{session_id}] Beolvasott journal sorok: {nrow(journaldata)}")
        
        incProgress(0.2, detail = "Profil adatok...")
        log_info("[{session_id}] MTMT profil lekérése: {mtmtid}")
        call_mtmt_prof <- paste0("https://m2.mtmt.hu/api/author/", mtmtid)
        Sys.sleep(0.1)
        api_result <- GET(call_mtmt_prof)
        if (status_code(api_result) != 200) {
          log_error("[{session_id}] API hiba (Profil): {status_code(api_result)}")
          res.json$message
          stop("Nem sikerült elérni az MTMT API-t.")
        }
        res.json <- fromJSON(content(api_result, "text", encoding="UTF-8"), flatten = TRUE)
        
        phd_info <- res.json$content$degrees$label
        phd_year <- as.numeric(stringr::str_extract(phd_info, "\\d{4}"))
        name <- paste(res.json$content$familyName, res.json$content$givenName)
        aux <- res.json$content$auxName
        affiliaciok <- res.json$content$affiliations %>% pull(worksFor.label)
        log_info("[{session_id}] Szerző azonosítva: {name}")
        
        incProgress(0.3, detail = "Publikációk letöltése...")
        log_info("[{session_id}] Publikációk lekérése")
        pub_url <- paste0("https://m2.mtmt.hu/api/publication?cond=authorships.author;eq;", mtmtid, "&size=5000&format=json")
        pub_res <- GET(pub_url)
        res.json_pub <- fromJSON(content(pub_res, "text", encoding="UTF-8"), flatten = TRUE)
        publications <- as_tibble(res.json_pub$content)
        log_info("[{session_id}] Letöltött publikációk száma: {nrow(publications)}")
        if(nrow(publications)==0) {
          log_error("[{session_id}] Letöltött publikációk száma: {nrow(publications)}")
          stop("A publikációknál nincs megjelenítendő elem.")
        }
        
        kozl_szama <- nrow(publications)
        tud_kozl_szama <- publications %>% subset(category.label == "Tudományos") %>% nrow()
        
        incProgress(0.2, detail = "Excel kitöltése...")
        log_info("[{session_id}] Excel INFO sheet kitöltése")
        
        wb <- loadWorkbook(output_path)
        
        writeData(wb, "INFO", x = mtmtid, startCol = 2, startRow = 1)
        writeData(wb, "INFO", x = name, startCol = 2, startRow = 2)
        writeData(wb, "INFO", x = aux, startCol = 2, startRow = 3)
        writeData(wb, "INFO", x = phd_year, startCol = 2, startRow = 5)
        writeData(wb, "INFO", x = kozl_szama, startCol = 2, startRow = 6)
        writeData(wb, "INFO", x = tud_kozl_szama, startCol = 2, startRow = 7)
        writeData(wb, "INFO", x = affiliaciok, startCol = 2, startRow = 8)
        
        apply_cond_format(wb, "INFO", cols = 2,  rows = 4, rule = "LEN(B4)=0")
        
        # Verziószám hozzáadása
        writeData(wb, "INFO", x = paste0("Jelenlegi verzió: ", aktualis_verzioszam, " (", aktualis_datum, ")"), startCol = 1, startRow = 8+length(affiliaciok))
        addStyle(wb,
                 sheet = "INFO",
                 style = createStyle(textDecoration = c("bold", "italic"), fontColour = "#808080"),
                 cols = 1, rows = 8+length(affiliaciok),
                 gridExpand = FALSE)
        
        # ---------------------------------------------------------
        # 3. Új változók létrehozása és meglévők átalakítása
        # ---------------------------------------------------------
        
        log_info("[{session_id}] Új változók létrehozása és meglévők átalakítása")
        
        # --- Lang változó
        publications <- publications %>%
          mutate(lang = map_chr(languages, function(df) {
            # Ellenőrizzük, hogy null-e vagy üres-e a data.frame
            if (is.null(df) || nrow(df) == 0) {
              return(NA_character_)
            }
            
            # Kinyerjük a 'label' oszlop első sorát
            # A pull() vagy a $ használható, a [1] garantálja az 1-es hosszt
            res <- df$label[1]
            
            return(as.character(res))
          }))
        
        # Ellenőrzés
        table(publications$lang, useNA = "always")
        
        # --- mtaRatingsForSort változó frissítése (1)
        # Frissítés a IX. Szociológiai Tudományos Bizottság értékelése alapján.
        # Azért szükséges, mert a json betöltésekor automatikusan az elsőt vette ki a listából,
        # de nekünk arra az elemre van szükségünk, ahol a ratingType.label == "IX. Szociológiai Tudományos Bizottság".
        
        if (!("ratings" %in% colnames(publications))) {
          publications$ratings <- NA_character_
        }
        
        publications <- publications %>%
          mutate(
            # Egy ideiglenes oszlopba kinyerünk mindent egyszerre
            extracted_ratings = map(ratings, ~{
              # Alapértelmezett kimenetek, ha a lista üres
              out <- list(rating_sztb = NA_character_, IX = NA_character_, norveg = FALSE)
              
              if (is.null(.x) || length(.x) == 0) return(out)
              df <- as.data.frame(.x)
              
              # SZTB és norveg kinyerése
              if (all(c("ratingType.code", "val") %in% names(df))) {
                sztb_val <- df$val[df$ratingType.code == "SZTB"]
                if (length(sztb_val) > 0) out$rating_sztb <- as.character(sztb_val[1])
                
                out$norveg <- any(df$ratingType.code == "norveg")
              }
              
              # IX kinyerése (label alapján)
              if (all(c("label", "val") %in% names(df))) {
                ix_idx <- which(startsWith(as.character(df$label), "IX"))
                if (length(ix_idx) > 0) {
                  res <- as.character(df$val[ix_idx[1]])
                  if (!is.na(res) && nchar(res) > 0) substr(res, 1, 1) <- "D"
                  out$IX <- res
                }
              }
              
              return(out)
            })
          ) %>%
          # A listát szétbontjuk külön oszlopokká (rating_sztb, IX, norveg)
          unnest_wider(extracted_ratings)
        
        # Felülírjuk, mert csak az SZTB besorolás számít nekünk
        if ("mtaRatingsForSort" %in% colnames(publications)) {
          publications$mtaRatingsForSort <- publications$rating_sztb
        } else {
          publications$mtaRatingsForSort <- NA_character_
        }
        
        # --- mtaRatingsForSort változó frissítése (2)
        # Ha nincs SZTB: Frissítés a IX. osztály valamelyik bizottságának értékelése alapján.
        publications$mtaRatingsForSort <- ifelse(is.na(publications$mtaRatingsForSort),
                                                 publications$IX,
                                                 publications$mtaRatingsForSort)
        
        
        # --- mtaRatingsForSort változó frissítése (3)
        # Journaldata infók rárakása
        
        if ("journal.mtid" %in% colnames(publications)) {
          publications <- journaldata %>%
            select(`MTMT azonosító`,Értékelés,nemzetkozi) %>%
            rename("journal.mtid"=`MTMT azonosító`) %>%
            left_join(publications,.,by="journal.mtid")
          
          # rating frissítése a journal data alapján: ha üres, onnan pótol
          table(publications$mtaRatingsForSort)
          sum(is.na(publications$mtaRatingsForSort))
          publications$mtaRatingsForSort <- ifelse(is.na(publications$mtaRatingsForSort),
                                                   publications$Értékelés,
                                                   publications$mtaRatingsForSort)
          sum(is.na(publications$mtaRatingsForSort))
        }
        
        
        if (all(is.na(publications$mtaRatingsForSort))) {
          log_warn("[{session_id}] Figyelem: Minden mtaRatingsForSort mező üres!")
        }
        
        # --- categ változó (A, B, C, D)
        publications$categ <- substr(publications$mtaRatingsForSort,0,1)
        
        
        # --- phd_utan_keletkezett_kozl
        publications <- publications %>%
          mutate(phd_utan_keletkezett_kozl = NA)
        
        
        # --- ISBN kód kinyerése
        if ("book.identifiers" %in% colnames(publications)) {
          publications$isbn_extracted <- sapply(
            publications$book.identifiers,
            function(x) {
              if (is.null(x) || is.null(x$label)) return(NA_character_)
              
              v <- unlist(x$label)   # flatten list element
              out <- str_extract(v, "ISBN:\\s*[0-9Xx-]+")
              
              out[!is.na(out)][1] %||% NA_character_
            }
          )
        }
        
        
        # --- ertekelesi_ter_kod létrehozása
        # Ennek az a lényege, hogy trackeljük, mely közlemények kaptak már pontot.
        # Ha már foglalkoztunk a közleménnyel és bekerült az output excelbe, legyen ott az értékelési terület kódja.
        publications$ertekelesi_ter_kod <- "0"
        
        
        # --- Szerzők száma
        # Sokszor 0, ilyenkor kiszámoljuk az authorships alapján
        publications <- publications %>%
          mutate(authorCount = if_else(authorCount == 0,
                                       map_int(authorships, nrow),
                                       authorCount))
        
        # --- Hiányzó oszlopok ellenőrzése és pótlása
        # szükséges oszlopok listája
        required_cols <- c(
          "label","mtid","otype","publishedYear","firstAuthor","title","subTitle","lang","pageLength","subType.label",
          "category.label","ossz_iv","ossz_iv_normalt","norveg","citationCount","phd_utan_keletkezett_kozl","ertekelesi_ter_kod",
          "authorCount","book.conferencePublication","category.label","categ","nemzetkozi","norveg","citationCount",
          "phd_utan_keletkezett_kozl","ertekelesi_ter_kod","authorCount","isbn_extracted","journal.pIssn","journal.eIssn","volume",
          "issue","firstPage","lastPage")
        
        missing_cols <- setdiff(required_cols, names(publications))
        
        if (length(missing_cols) > 0) {
          # Létrehozunk egy listát, ahol minden hiányzó oszlop neve NA értéket kap
          # Ezt tibble-é alakítjuk
          new_cols <- missing_cols %>% 
            map_dfc(~ tibble(!!.x := NA))
          
          # Hozzáadjuk a hiányzó oszlopokat (bind_cols csak akkor jó, ha sorok száma egyezik, 
          # ezért inkább a mutate-tel injektálunk egy üres listát)
          publications <- publications %>%
            mutate(!!!setNames(rep(list(NA), length(missing_cols)), missing_cols))
        }
        
        
        # ---------------------------------------------------------
        # 4. Új sheet létrehozása a legördülő lista T/F értékeihez
        # ---------------------------------------------------------
        
        # Segéd munkalap a TRUE/FALSE értékekhez
        if (!("Lists" %in% names(wb))) {
          addWorksheet(wb, "Lists")
        }
        
        # Kiírjuk a TRUE/FALSE értékeket erre a lapra
        writeData(wb, "Lists", x = c(TRUE, FALSE), startCol = 1, startRow = 1)
        
        # Elrejtés
        sheetVisibility(wb)[which(names(wb) == "Lists")] <- FALSE
        
        
        # ---------------------------------------------------------
        # 5. Hivatkozások táblázat
        # ---------------------------------------------------------
        
        log_info("[{session_id}] Hivatkozások lekérése minden közleményhez")
        
        # --- Hivatkozásokat tartalmazó dataframe előállítása
        # Minden közlemény MTID azonosítójához lekérjük API-n keresztül az adatokat
        all_citations <- publications %>%
          rowwise() %>%
          mutate(
            data = list({
              # Use mtid from folyoirat
              this_mtid <- mtid
              
              call_mtmt <- paste0(
                "https://m2.mtmt.hu/api/publication?",
                "sort=publishedYear,desc&sort=firstAuthor&sort=title&size=100&",
                "cond=published;eq;true&",
                "cond=cites.publication;eq;", this_mtid, "&",
                "cond=cites.published;eq;true"
              )
              
              api_result <- GET(call_mtmt)
              if (status_code(api_result) != 200) return(NULL)
              
              json_result <- content(api_result, "text", encoding = "UTF-8")
              res.json <- fromJSON(json_result, flatten = TRUE)
              
              if (is.null(res.json$content)) return(NULL)
              
              # Convert to tibble and add original mtid from folyoirat
              as_tibble(res.json$content) %>%
                mutate(mtmtid_kozl = this_mtid, .before = 1)
            })
          ) %>%
          ungroup()
        
        # Hivatkozás-szintre hozás
        all_citations <- all_citations$data[!sapply(all_citations$data, is.null)] %>%
          bind_rows()
        
        if (!"predatorRatingsForSort" %in% names(all_citations)) {
          all_citations$predatorRatingsForSort <- 0
        }
        
        all_citations$predatorRatingsForSort <- as.integer(all_citations$predatorRatingsForSort)
        
        # --- Hiányzó oszlopok ellenőrzése és pótlása
        # szükséges oszlopok listája
        required_cols <- c("mtmtid_kozl","otype","predatorRatingsForSort","mtid","journal.mtid","label","ratings",
                           "mtaRatingsForSort","authorships","independentCitingPubCount","created","publishedYear")
        
        missing_cols <- setdiff(required_cols, names(all_citations))
        
        if (length(missing_cols) > 0) {
          # Létrehozunk egy listát, ahol minden hiányzó oszlop neve NA értéket kap
          # Ezt tibble-é alakítjuk
          new_cols <- missing_cols %>% 
            map_dfc(~ tibble(!!.x := NA))
          
          # Hozzáadjuk a hiányzó oszlopokat (bind_cols csak akkor jó, ha sorok száma egyezik, 
          # ezért inkább a mutate-tel injektálunk egy üres listát)
          all_citations <- all_citations %>%
            mutate(!!!setNames(rep(list(NA), length(missing_cols)), missing_cols))
        }
        
        # Oszlopok kiválasztása
        all_citations <- all_citations %>%
          select(mtmtid_kozl,otype,predatorRatingsForSort,mtid,journal.mtid,label,ratings,mtaRatingsForSort,authorships,independentCitingPubCount,created,publishedYear) %>%
          rename("norveg"=predatorRatingsForSort)
        all_citations$norveg[which(is.na(all_citations$norveg)==F & all_citations$norveg!=0)] <- 1
        all_citations$norveg[which(is.na(all_citations$norveg))] <- 0
        all_citations$created <- as.Date(substr(all_citations$created,1,10))
        
        
        # --- Self citation detektálása
        
        # Szedjük ki a publikációk szerzőit
        authors_of_publication <- publications %>% select(mtid, authorships, label)
        
        # Tegyük rá a publikációk szerzőit arra a dataframe-ra, amiben a hivatkozások és azok szerzői vannak
        all_citations <- authors_of_publication %>%
          select(-c(label)) %>%
          rename("authors_of_publ"=authorships,
                 "mtmtid_kozl"=mtid) %>%
          left_join(all_citations, ., by="mtmtid_kozl")
        
        # self_citations detection
        all_citations$self_citation <- mapply(
          function(publ, auth) {
            
            # Extract author.mtid safely from authors_of_publ
            publ_ids <- if (!is.null(publ) && "author.mtid" %in% names(publ)) {
              publ$author.mtid
            } else {
              character(0)
            }
            
            # Extract author.mtid safely from authorships
            auth_ids <- if (!is.null(auth) && "author.mtid" %in% names(auth)) {
              auth$author.mtid
            } else {
              character(0)
            }
            
            # Flag = 1 if any overlap, else 0
            as.integer(length(intersect(publ_ids, auth_ids)) > 0)
          },
          
          all_citations$authors_of_publ,
          all_citations$authorships
        )
        
        
        # --- mtaRatingsForSort változó frissítése:
        # először a IX. Szociológiai Tudományos Bizottság értékelése alapján
        # majd a IX. osztály valamelyik bizottság értékelése alapján
        # van benne hazai / nemzetközi információ?
        
        if (!("ratings" %in% colnames(all_citations))) {
          all_citations$ratings <- NA_character_
        }
        
        all_citations <- all_citations %>%
          mutate(
            # 1. Kinyerjük az adatokat a listákból egyetlen menetben
            extracted_ratings = map(ratings, ~{
              out <- list(rating_sztb = NA_character_, IX = NA_character_, hazai_nemzetkozi = NA_character_)
              
              if (is.null(.x) || length(.x) == 0) return(out)
              df <- as.data.frame(.x)
              
              # SZTB kinyerése
              if (all(c("ratingType.code", "val") %in% names(df))) {
                sztb_val <- df$val[df$ratingType.code == "SZTB"]
                if (length(sztb_val) > 0) out$rating_sztb <- as.character(sztb_val[1])
              }
              
              # IX kinyerése (label alapján)
              if (all(c("label", "val") %in% names(df))) {
                ix_idx <- which(startsWith(as.character(df$label), "IX"))
                if (length(ix_idx) > 0) {
                  res <- as.character(df$val[ix_idx[1]])
                  if (!is.na(res) && nchar(res) > 0) substr(res, 1, 1) <- "D"
                  out$IX <- res
                }
              }
              
              # Hazai/nemzetközi infó kinyerése (val alapján)
              if ("val" %in% names(df)) {
                hn_idx <- which(stringr::str_detect(as.character(df$val), "hazai|nemzetközi"))
                if (length(hn_idx) > 0) {
                  res_hn <- as.character(df$val[hn_idx[1]])
                  # A 3. karaktertől vesszük a szöveget
                  if (!is.na(res_hn) && nchar(res_hn) >= 3) {
                    out$hazai_nemzetkozi <- substr(res_hn, 3, nchar(res_hn))
                  }
                }
              }
              
              return(out)
            })
          ) %>%
          # Szétszedjük a listát oszlopokra
          unnest_wider(extracted_ratings)
        
        if (!("rating_sztb" %in% colnames(all_citations))) {
          all_citations$rating_sztb <- NA_character_
        }
        
        if (!("IX" %in% colnames(all_citations))) {
          all_citations$IX <- NA_character_
        }
        
        if (!("hazai_nemzetkozi" %in% colnames(all_citations))) {
          all_citations$hazai_nemzetkozi <- NA_character_
        }
            
        # Értékek konszolidálása (az eredeti ifelse-ek kiváltása)
        all_citations <- all_citations %>% 
          mutate(
            # A coalesce() az első nem NA értéket veszi (először SZTB, ha nincs, akkor IX)
            mtaRatingsForSort = coalesce(rating_sztb, IX),
            
            # Ha a minősítés csak 1 karakter hosszú (pl. "D"), és van hazai/nemz infó, összefűzzük
            mtaRatingsForSort = if_else(
              nchar(mtaRatingsForSort) == 1 & !is.na(hazai_nemzetkozi),
              paste(mtaRatingsForSort, hazai_nemzetkozi),
              mtaRatingsForSort
            )
          ) %>%
          # Eldobjuk a feleslegessé vált segédoszlopokat
          select(-rating_sztb, -IX, -hazai_nemzetkozi)

        
        # --- mtaRatingsForSort változó frissítése a journaldata alapján
        # Journaldata infók rárakása
        all_citations <- journaldata %>%
          select(`MTMT azonosító`,Értékelés,nemzetkozi) %>%
          rename("journal.mtid"=`MTMT azonosító`) %>%
          left_join(all_citations,.,by="journal.mtid")
        
        # rating frissítése
        all_citations$mtaRatingsForSort <- ifelse(is.na(all_citations$mtaRatingsForSort),
                                                  all_citations$Értékelés,
                                                  all_citations$mtaRatingsForSort)
        
        if (all(is.na(all_citations$mtaRatingsForSort))) {
          log_warn("[{session_id}] Figyelem: Minden mtaRatingsForSort mező üres!")
        }
        
        all_citations$categ_hiv <- "további teljes tudományos közleményben"
        all_citations$categ_hiv[which(all_citations$nemzetkozi==T | str_detect(all_citations$mtaRatingsForSort, "nemzetközi"))] <- "nemzetközi listás folyóiratban"
        all_citations$categ_hiv[which(all_citations$nemzetkozi==F | str_detect(all_citations$mtaRatingsForSort, "hazai"))] <- "nem nemzetközi listás folyóiratban"
        table(all_citations$categ_hiv)
        colnames(all_citations)
        
        all_citations <- all_citations %>%
          select(mtmtid_kozl,mtid,norveg,otype,journal.mtid,label,mtaRatingsForSort,nemzetkozi,categ_hiv,independentCitingPubCount,created,publishedYear,self_citation)
        
        all_citations <- publications %>%
          select(mtid,label,category.label) %>%
          rename("mtmtid_kozl"=mtid,
                 "label_kozl"=label,
                 "category.label_kozl"=category.label) %>% # Rárakni, hogy a publikáció tudományos-e
          left_join(all_citations,.,by="mtmtid_kozl") %>%
          select(label_kozl,mtmtid_kozl,category.label_kozl,mtid,otype,journal.mtid,label,mtaRatingsForSort,norveg,categ_hiv:self_citation) %>%
          mutate(
            date_min = pmin(
              created,
              as.Date(sprintf("%04d-01-01", publishedYear))
            )
          )
        
        if (nrow(all_citations)==0) {
          all_citations$mtmtid_kozl <- as.character(all_citations$mtmtid_kozl)
          all_citations[1, ] <- NA          # create one empty row
          all_citations[1, 1] <- "A jelölt közleményei nem rendelkeznek hivatkozással."
        }
        
        all_citations <- all_citations %>%
          rename(
            "Közlemény" = "label_kozl",
            "Közlemény mtid" = "mtmtid_kozl",
            "Közlemény Kategória" = "category.label_kozl",
            "mtid" = "mtid",
            "Típus" = "otype",
            "Lap mtid" = "journal.mtid",
            "Hivatkozás" = "label",
            "MTA Besorolás" = "mtaRatingsForSort",
            "Norvég lista" = "norveg",
            "Lap értékelése kategória" = "categ_hiv",
            "Független hivatkozások száma" = "independentCitingPubCount",
            "Létrehozás" = "created",
            "Publikálás éve" = "publishedYear",
            "Önhivatkozás" = "self_citation",
            "Dátum minimum" = "date_min"
          )
        
        # mentés lapfülre
        writeData(wb, "4_hivatkozasok", x = all_citations, startCol = 1, startRow = 25)
        
        
        # --- Hiperhivatkozások
        
        linkek <- paste0("https://m2.mtmt.hu/api/publication/", all_citations$`Közlemény mtid`)
        
        # Szöveg tisztítása (Idézőjelek lecserélése aposztrófra, hogy ne törje meg a formulát)
        megjelenitett_szoveg <- gsub('"', "'", all_citations[[1]])
        
        # Formulák összeállítása
        # Fontos: Az Excelben a pontosvessző vagy vessző elválasztó az oprendszer nyelvétől függhet,
        # de az openxlsx a vesszőt várja (angol szintaxis), amit az Excel fordít le.
        formulas <- paste0('HYPERLINK("', linkek, '", "', megjelenitett_szoveg, '")')
        
        # Írás a munkalapra
        writeFormula(wb, sheet = "4_hivatkozasok", x = formulas, startCol = 1, startRow = 26)
        
        # Stílus alkalmazása
        link_style <- createStyle(fontColour = "#0000FF", textDecoration = "underline")
        addStyle(wb, sheet = "4_hivatkozasok",
                 style = link_style,
                 rows = 26:(25 + nrow(all_citations)),
                 cols = 1,
                 gridExpand = TRUE)
        
        rm(linkek, megjelenitett_szoveg, formulas, link_style)
        
        
        # --- PhD fokozat megszerzését követően írt közlemény
        
        # 1. A sablon (Ez maradhat változatlan)
        formula_template <- paste0(
          '=IFERROR(',
          'VLOOKUP(B%d, \'2_folyoirat\'!$B:$T, 19, FALSE), ',
          'IFERROR(',
          'VLOOKUP(B%d, \'1_szakkonyv_monog\'!$B:$P, 15, FALSE), ',
          'VLOOKUP(B%d, \'3_tovabbi\'!$B:$Q, 16, FALSE)',
          ')',
          ')'
        )
        
        # 2. Sorok meghatározása
        rows <- 26:(26+nrow(all_citations)-1)
        
        # 3. Képletek generálása - JAVÍTOTT RÉSZ
        # A gsub lecseréli az összes "%d"-t az aktuális 'r' sorszámra.
        # Nem kell str_count, rep vagy do.call.
        formulas <- sapply(rows, function(r) {
          gsub("%d", r, formula_template, fixed = TRUE)
        })
        
        writeFormula(wb, "4_hivatkozasok", x = formulas, startCol = 16, startRow = 26)
        
        
        # --- Ér pontot a hivatkozás?
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=AND(N%d=0, I%d=0, OR(E%d="JournalArticle", E%d="Book", E%d="BookChapter"))'
        )
        
        rows <- 26:(26+nrow(all_citations)-1)
        n_placeholders <- stringr::str_count(formula_template, "%d")
        formulas <- sapply(rows, function(r) {
          do.call(sprintf, c(list(fmt = formula_template), as.list(rep(r, n_placeholders))))
        })
        
        writeFormula(wb, "4_hivatkozasok", x = formulas, startCol = 17, startRow = 26)
        
        
        # --- Ér pontot a hivatkozás, mint a fokozat megszerzését követően írt közleményre történt hivatkozás? (Szempont="a PhD utáni közleményekre")
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=AND(N%d=0, I%d=0, P%d=TRUE, OR(E%d="JournalArticle", E%d="Book", E%d="BookChapter"))'
        )
        
        rows <- 26:(26+nrow(all_citations)-1)
        n_placeholders <- stringr::str_count(formula_template, "%d")
        formulas <- sapply(rows, function(r) {
          do.call(sprintf, c(list(fmt = formula_template), as.list(rep(r, n_placeholders))))
        })
        
        writeFormula(wb, "4_hivatkozasok", x = formulas, startCol = 18, startRow = 26)
        
        
        # --- Ér pontot a hivatkozás, mint a fokozat megszerzését követően történt hivatkozás? (Szempont="a PhD óta")
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=AND(N%d=0, I%d=0, O%d>(INFO!$B$4), OR(E%d="JournalArticle", E%d="Book", E%d="BookChapter"))'
        )
        
        rows <- 26:(26+nrow(all_citations)-1)
        n_placeholders <- stringr::str_count(formula_template, "%d")
        formulas <- sapply(rows, function(r) {
          do.call(sprintf, c(list(fmt = formula_template), as.list(rep(r, n_placeholders))))
        })
        
        writeFormula(wb, "4_hivatkozasok", x = formulas, startCol = 19, startRow = 26)
        
        
        # --- I oszlop szám legyen
        
        #szam_stilus <- createStyle(numFmt = "#,##0")
        #addStyle(
        #  wb,
        #  sheet = "4_hivatkozasok",
        #  style = szam_stilus,
        #  rows = 26:(26+nrow(all_citations)-1),
        #  cols = 9,
        #  gridExpand = TRUE,
        #  stack = TRUE
        #)

        
        # --- Feltételes formázások
        
        rows_hiv <- 26:(26+nrow(all_citations)-1)
        apply_cond_format(wb, "4_hivatkozasok", cols = 10,  rows = rows_hiv, rule = "LEN(J26)=0")
        apply_cond_format(wb, "4_hivatkozasok", cols = 13,  rows = rows_hiv, rule = "M26=INFO!$B$5")
        apply_cond_format(wb, "4_hivatkozasok", cols = 15,  rows = rows_hiv, rule = "YEAR(O26)=INFO!$B$5")

        
        # --- További feltételes formázás
        if (sum(all_citations$Közlemény=="A jelölt közleményei nem rendelkeznek hivatkozással.")!=0) {
          addStyle(
            wb,
            sheet = "4_hivatkozasok",
            style = createStyle(textDecoration = c("bold", "italic"), fontColour = "#FF0000"),
            cols = 1, rows = 26,
            gridExpand = FALSE)
        }
        
        
        # ---------------------------------------------------------
        # 6. Output tábla feltöltése
        # (1) Kiemelten értékelt szakkönyv, monográfia írása (legfeljebb 3 db, 112 oldal fölötti könyv az előző fokozat után)
        # ---------------------------------------------------------
        
        log_info("[{session_id}] (1) Kiemelten értékelt szakkönyv, monográfia írása")
        
        # Probléma: pageLength hiányos (6 könyvből 3-nál van), firstPage és lastPage NA-k, nem lehet máshonnan kiszedni az infót
        # Emiatt a pageLengthre nem szűrök
        # "az előző fokozat után" = PhD után
        
        # --- Közlemények táblázat
        
        # Ha Book, akkor legyen 1, de ha nagyobb vagy egyenlő a szerzők száma mint 2, akkor automatikusan 3.1
        publications$ertekelesi_ter_kod[which(publications$otype=="Book")] <- "1"
        publications$ertekelesi_ter_kod[which(publications$otype=="Book" & publications$authorCount > 2)] <- "3.1"
        # a szekció végén ezeket a könyveket 3.1-re rakni, hogy a 3_tovabbi lapfülön is megjelenjenek.
        # végül a jelölt dönt majd, melyik hova számítódjon, de ha valahol a beszámítás 1, akkor a másik helyen 0 legyen.
        
        konyv <- publications %>%
          subset(otype=="Book") %>%
          subset(authorCount <= 2) %>%  #legfeljebb egy társszerzős
          subset(ertekelesi_ter_kod=="1") %>% # TODO: ezt kikommentelni, ha tesztelni akarunk!
          mutate(ossz_iv = NA) %>%
          mutate(ossz_iv_normalt = NA) %>%
          select(label,mtid,otype,publishedYear,firstAuthor,title,subTitle,lang,pageLength,subType.label,category.label,ossz_iv,ossz_iv_normalt,norveg,citationCount,phd_utan_keletkezett_kozl,ertekelesi_ter_kod,authorCount)
        
        # Itt át is állítjuk az értékelési terület kódot, hogy a könyvek a 3.1-be is bekerüljenek,
        # annak érdekében, hogy választani lehessen, hova számolja le a jelölt.
        publications$ertekelesi_ter_kod[which(publications$otype=="Book")] <- "3.1"
        
        if (nrow(konyv)==0) {
          konyv[1, ] <- NA          # create one empty row
          konyv[1, 1] <- "A jelölt nem rendelkezik kiemelten értékelt szakkönyvvel, monográfiával."
        }
        
        # Új oszlopok hozzáadása, amik az output excelben is vannak:
        konyv <- konyv %>% mutate(ív=NA, .before = "ertekelesi_ter_kod")
        konyv <- konyv %>% mutate(pontszám=NA, .before = "ertekelesi_ter_kod")
        
        konyv <- konyv %>%
          mutate(hiv_nemz_list_folyoirat=NA,
                 hiv_nem_nemz_list_folyoirat=NA,
                 hiv_tov_telj_tud_kozlemeny=NA,
                 hiv_nemz_list_folyoirat_pontszám=NA,
                 hiv_nem_nemz_list_folyoirat_pontszám=NA,
                 hiv_tov_telj_tud_kozlemeny_pontszám=NA,
                 citationCount_eq_hiv=NA)
        
        konyv <- konyv %>%
          rename(
            "Közlemény" = "label",
            "mtid" = "mtid",
            "Típus" = "otype",
            "Publikálás éve" = "publishedYear",
            "Első szerző" = "firstAuthor",
            "Cím" = "title",
            "Alcím" = "subTitle",
            "Nyelv" = "lang",
            "Oldalszám" = "pageLength",
            "Altípus" = "subType.label",
            "Kategória" = "category.label",
            "Össz-ív" = "ossz_iv",
            "Szerzők számával normált ívszám" = "ossz_iv_normalt",
            "Norvég lista" = "norveg",
            "Citációk száma" = "citationCount",
            "PhD után keletkezett közlemény" = "phd_utan_keletkezett_kozl",
            "Saját ívek száma" = "ív",
            "Pontszám" = "pontszám",
            "Értékelési terület kódja" = "ertekelesi_ter_kod",
            "Hivatkozás nemzetközi listás folyóiratban (db)" = "hiv_nemz_list_folyoirat",
            "Hivatkozás nem nemzetközi listás folyóiratban (db)" = "hiv_nem_nemz_list_folyoirat",
            "Hivatkozás további teljes tudományos közleményben (db)" = "hiv_tov_telj_tud_kozlemeny",
            "Hivatkozás nemzetközi listás folyóiratban (pontszám)" = "hiv_nemz_list_folyoirat_pontszám",
            "Hivatkozás nem nemzetközi listás folyóiratban (pontszám)" = "hiv_nem_nemz_list_folyoirat_pontszám",
            "Hivatkozás további teljes tudományos közleményben (pontszám)" = "hiv_tov_telj_tud_kozlemeny_pontszám",
            "Citációk száma egyezik-e a hivatkozások számával" = "citationCount_eq_hiv",
            "Szerzők száma" = "authorCount"
          ) %>%
          mutate("Hivatkozás a phd fokozat megszerzése óta nemzetközi listás folyóiratban (db)" = NA,
                 "Hivatkozás a phd fokozat megszerzése óta nem nemzetközi listás folyóiratban (db)" = NA,
                 "Hivatkozás a phd fokozat megszerzése óta további teljes tudományos közleményben (db)" = NA,
                 "Hivatkozás a phd fokozat megszerzése óta nemzetközi listás folyóiratban (pontszám)" = NA,
                 "Hivatkozás a phd fokozat megszerzése óta nem nemzetközi listás folyóiratban (pontszám)" = NA,
                 "Hivatkozás a phd fokozat megszerzése óta további teljes tudományos közleményben (pontszám)" = NA) %>%
          mutate(Pontbeszámítás = "Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám") %>%
          relocate("Szerzők száma", .after = "Pontbeszámítás") %>%
          mutate("Könyv után járó pontszám számolása" = "Saját ívek száma alapján")
        
        writeData(wb, "1_szakkonyv_monog", x = konyv, startCol = 1, startRow = 12)
        
        
        # --- Hiperhivatkozások
        
        add_mtmt_hyperlinks(wb, "1_szakkonyv_monog", df = konyv, id_col = "mtid", start_row = 13)
        
        
        # --- phd_utan_keletkezett_kozl függvény
        
        formula_template <- paste0(
          '=D%d>(INFO!$B$5)'
        )
        rows <- 13:(13+nrow(konyv)-1)
        write_dynamic_formula(wb, "1_szakkonyv_monog", formula_template, rows, start_col = 16, start_row = 13)
        

        # --- Szerzők számával normált ívszám
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(L%d="", "", L%d/AH%d)'
        )
        rows <- 13:(13+nrow(konyv)-1)
        write_dynamic_formula(wb, "1_szakkonyv_monog", formula_template, rows, start_col = 13, start_row = 13)
        
        
        # --- Közlemények után járó pontszám képlet
        
        # Függvény dinamikus elkészítése az összes sorra (társszerzős szorzó nélkül)
        formula_template <- paste0(
          '=IF(AI%d="Saját ívek száma alapján", ',
          # EREDETI FÜGGVÉNY (Q oszloppal)
          'IF(OR(H%d="", Q%d="", P%d="", AG%d="", AH%d=""), "", ',
          'IF(AG%d<>"Közlemény és hivatkozások után járó pontszám beszámítása", 0, ',
          'IF(P%d=FALSE, 0, ',
          'IF(N%d=TRUE, 0, ',
          'IF(I%d<=112, 0, ',
          'IF(Q%d<=7, 0, ',
          'IF(AND(H%d<>"Magyar", P%d=TRUE), Q%d*4, ',
          'IF(AND(H%d="Magyar", P%d=TRUE), Q%d*2)',
          '))))))), ',
          # HA NEM "Saját ívek száma alapján" (M oszloppal)
          'IF(OR(H%d="", M%d="", P%d="", AG%d="", AH%d=""), "", ',
          'IF(AG%d<>"Közlemény és hivatkozások után járó pontszám beszámítása", 0, ',
          'IF(P%d=FALSE, 0, ',
          'IF(N%d=TRUE, 0, ',
          'IF(I%d<=112, 0, ',
          'IF(M%d<=7, 0, ',
          'IF(AND(H%d<>"Magyar", P%d=TRUE), M%d*4, ',
          'IF(AND(H%d="Magyar", P%d=TRUE), M%d*2)',
          ')))))))',
          ')' # Legkülső IF lezárása
        )
        
        rows <- 13:(13+nrow(konyv)-1)
        write_dynamic_formula(wb, "1_szakkonyv_monog", formula_template, rows, start_col = 18, start_row = 13)

        
        # --- Hivatkozások száma - hiv_nemz_list_folyoirat
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "nemzetközi listás folyóiratban", ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 13:(13+nrow(konyv)-1)
        write_dynamic_formula(wb, "1_szakkonyv_monog", formula_template, rows, start_col = 20, start_row = 13)
        
        # --- Hivatkozások száma - hiv_nem_nemz_list_folyoirat
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "nem nemzetközi listás folyóiratban", ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 13:(13+nrow(konyv)-1)
        write_dynamic_formula(wb, "1_szakkonyv_monog", formula_template, rows, start_col = 21, start_row = 13)
        
        
        # --- Hivatkozások száma - hiv_tov_telj_tud_kozlemeny
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "további teljes tudományos közleményben", ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 13:(13+nrow(konyv)-1)
        write_dynamic_formula(wb, "1_szakkonyv_monog", formula_template, rows, start_col = 22, start_row = 13)
        
        
        # --- Hivatkozások pontszám: hiv_nemz_list_folyoirat_pontszám
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(AI%d="Saját ívek száma alapján", ',
          # 1. ÁG: "Saját ívek száma alapján" -> Mindenhol Q%d szerepel
          'IF(OR(H%d="", I%d="", Q%d="", T%d="", U%d="", V%d="", AG%d="", AH%d=""), "", ',
          'IF(AG%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(N%d=TRUE, 0, ',
          'IF(I%d<=112, 0, ',
          'IF(Q%d<=7, 0, ',
          'IF(H%d<>"Magyar", T%d*0.6*Q%d, T%d*0.3*Q%d) * (Q%d/L%d)',
          '))))), ',
          # 2. ÁG: Minden más eset (Szerzők számával normált ívszám) -> Q helyett M%d szerepel
          'IF(OR(H%d="", I%d="", M%d="", T%d="", U%d="", V%d="", AG%d="", AH%d=""), "", ',
          'IF(AG%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(N%d=TRUE, 0, ',
          'IF(I%d<=112, 0, ',
          'IF(M%d<=7, 0, ',
          'IF(H%d<>"Magyar", T%d*0.6*M%d, T%d*0.3*M%d) * (M%d/L%d)',
          ')))))',
          ')' # Legkülső IF lezárása
        )
        
        rows <- 13:(13+nrow(konyv)-1)
        write_dynamic_formula(wb, "1_szakkonyv_monog", formula_template, rows, start_col = 23, start_row = 13)
      
        
        # --- Hivatkozások pontszám: hiv_nem_nemz_list_folyoirat_pontszám
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(AI%d="Saját ívek száma alapján", ',
          # 1. ÁG: "Saját ívek száma alapján" -> Mindenhol Q%d szerepel
          'IF(OR(H%d="", I%d="", Q%d="", T%d="", U%d="", V%d="", AG%d="", AH%d=""), "", ',
          'IF(AG%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(N%d=TRUE, 0, ',
          'IF(I%d<=112, 0, ',
          'IF(Q%d<=7, 0, ',
          'IF(H%d<>"Magyar", U%d*0.3*Q%d, U%d*0.15*Q%d) * (Q%d/L%d)',
          '))))), ',
          # 2. ÁG: Minden más eset (Szerzők számával normált ívszám) -> Q helyett M%d szerepel
          'IF(OR(H%d="", I%d="", M%d="", T%d="", U%d="", V%d="", AG%d="", AH%d=""), "", ',
          'IF(AG%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(N%d=TRUE, 0, ',
          'IF(I%d<=112, 0, ',
          'IF(M%d<=7, 0, ',
          'IF(H%d<>"Magyar", U%d*0.3*M%d, U%d*0.15*M%d) * (M%d/L%d)',
          ')))))',
          ')' # Legkülső IF lezárása
        )
        
        rows <- 13:(13+nrow(konyv)-1)
        write_dynamic_formula(wb, "1_szakkonyv_monog", formula_template, rows, start_col = 24, start_row = 13)
      
        
        # --- Hivatkozások pontszám: hiv_tov_telj_tud_kozlemeny_pontszám
        
        formula_template <- paste0(
          '=IF(AI%d="Saját ívek száma alapján", ',
          # 1. ÁG: "Saját ívek száma alapján" -> Mindenhol Q%d szerepel
          'IF(OR(H%d="", I%d="", Q%d="", T%d="", U%d="", V%d="", AG%d="", AH%d=""), "", ',
          'IF(AG%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(N%d=TRUE, 0, ',
          'IF(I%d<=112, 0, ',
          'IF(Q%d<=7, 0, ',
          'IF(H%d<>"Magyar", V%d*0.15*Q%d, V%d*0.1*Q%d) * (Q%d/L%d)',
          '))))), ',
          # 2. ÁG: Minden más eset (Szerzők számával normált ívszám) -> Q helyett M%d szerepel
          'IF(OR(H%d="", I%d="", M%d="", T%d="", U%d="", V%d="", AG%d="", AH%d=""), "", ',
          'IF(AG%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(N%d=TRUE, 0, ',
          'IF(I%d<=112, 0, ',
          'IF(M%d<=7, 0, ',
          'IF(H%d<>"Magyar", V%d*0.15*M%d, V%d*0.1*M%d) * (M%d/L%d)',
          ')))))',
          ')' # Legkülső IF lezárása
        )
        
        
        rows <- 13:(13+nrow(konyv)-1)
        write_dynamic_formula(wb, "1_szakkonyv_monog", formula_template, rows, start_col = 25, start_row = 13)
        
        
        # --- citationCount_eq_hiv
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=O%d=(T%d+U%d+V%d)'
        )
        
        rows <- 13:(13+nrow(konyv)-1)
        write_dynamic_formula(wb, "1_szakkonyv_monog", formula_template, rows, start_col = 26, start_row = 13)
        
        # --- Hivatkozások száma a phd megszerzése óta - hiv_phd_ota_nemz_list_folyoirat
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "nemzetközi listás folyóiratban", ',
          '4_hivatkozasok!$O$25:O10000, ">"&INFO!$B$4, ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 13:(13+nrow(konyv)-1)
        write_dynamic_formula(wb, "1_szakkonyv_monog", formula_template, rows, start_col = 27, start_row = 13)
        
        
        # --- Hivatkozások száma a phd megszerzése óta - hiv_nem_nemz_list_folyoirat
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "nem nemzetközi listás folyóiratban", ',
          '4_hivatkozasok!$O$25:O10000, ">"&INFO!$B$4, ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 13:(13+nrow(konyv)-1)
        write_dynamic_formula(wb, "1_szakkonyv_monog", formula_template, rows, start_col = 28, start_row = 13)
        
        
        # --- Hivatkozások száma a phd megszerzése óta - hiv_tov_telj_tud_kozlemeny
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "további teljes tudományos közleményben", ',
          '4_hivatkozasok!$O$25:O10000, ">"&INFO!$B$4, ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 13:(13+nrow(konyv)-1)
        write_dynamic_formula(wb, "1_szakkonyv_monog", formula_template, rows, start_col = 29, start_row = 13)
        
        
        # --- Hivatkozások pontszám a phd megszerzése óta: hiv_phd_ota_nemz_list_folyoirat_pontszám
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(AI%d="Saját ívek száma alapján", ',
          # 1. ÁG: "Saját ívek száma alapján" -> Mindenhol Q%d szerepel
          'IF(OR(H%d="", I%d="", Q%d="", AA%d="", AB%d="", AC%d="", AG%d="", AH%d=""), "", ',
          'IF(AG%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(N%d=TRUE, 0, ',
          'IF(I%d<=112, 0, ',
          'IF(Q%d<=7, 0, ',
          'IF(H%d<>"Magyar", AA%d*0.6*Q%d, AA%d*0.3*Q%d) * (Q%d/L%d)',
          '))))), ',
          # 2. ÁG: Minden más eset (Szerzők számával normált ívszám) -> Q helyett M%d szerepel
          'IF(OR(H%d="", I%d="", M%d="", AA%d="", AB%d="", AC%d="", AG%d="", AH%d=""), "", ',
          'IF(AG%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(N%d=TRUE, 0, ',
          'IF(I%d<=112, 0, ',
          'IF(M%d<=7, 0, ',
          'IF(H%d<>"Magyar", AA%d*0.6*M%d, AA%d*0.3*M%d) * (M%d/L%d)',
          ')))))',
          ')' # Legkülső IF lezárása
        )
        
        rows <- 13:(13+nrow(konyv)-1)
        write_dynamic_formula(wb, "1_szakkonyv_monog", formula_template, rows, start_col = 30, start_row = 13)
      
        
        # --- Hivatkozások pontszám a phd megszerzése óta: hiv_pdh_ota_nem_nemz_list_folyoirat_pontszám
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(AI%d="Saját ívek száma alapján", ',
          # 1. ÁG: "Saját ívek száma alapján" -> Mindenhol Q%d szerepel
          'IF(OR(H%d="", I%d="", Q%d="", AA%d="", AB%d="", AC%d="", AG%d="", AH%d=""), "", ',
          'IF(AG%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(N%d=TRUE, 0, ',
          'IF(I%d<=112, 0, ',
          'IF(Q%d<=7, 0, ',
          'IF(H%d<>"Magyar", AB%d*0.3*Q%d, AB%d*0.15*Q%d) * (Q%d/L%d)',
          '))))), ',
          # 2. ÁG: Minden más eset (Szerzők számával normált ívszám) -> Q helyett M%d szerepel
          'IF(OR(H%d="", I%d="", M%d="", AA%d="", AB%d="", AC%d="", AG%d="", AH%d=""), "", ',
          'IF(AG%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(N%d=TRUE, 0, ',
          'IF(I%d<=112, 0, ',
          'IF(M%d<=7, 0, ',
          'IF(H%d<>"Magyar", AB%d*0.3*M%d, AB%d*0.15*M%d) * (M%d/L%d)',
          ')))))',
          ')' # Legkülső IF lezárása
        )
        
        rows <- 13:(13+nrow(konyv)-1)
        write_dynamic_formula(wb, "1_szakkonyv_monog", formula_template, rows, start_col = 31, start_row = 13)
      
        
        # --- Hivatkozások pontszám a phd megszerzése óta: hiv_tov_telj_tud_kozlemeny_pontszám
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(AI%d="Saját ívek száma alapján", ',
          # 1. ÁG: "Saját ívek száma alapján" -> Mindenhol Q%d szerepel
          'IF(OR(H%d="", I%d="", Q%d="", AA%d="", AB%d="", AC%d="", AG%d="", AH%d=""), "", ',
          'IF(AG%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(N%d=TRUE, 0, ',
          'IF(I%d<=112, 0, ',
          'IF(Q%d<=7, 0, ',
          'IF(H%d<>"Magyar", AC%d*0.15*Q%d, AC%d*0.1*Q%d) * (Q%d/L%d)',
          '))))), ',
          # 2. ÁG: Minden más eset (Szerzők számával normált ívszám) -> Q helyett M%d szerepel
          'IF(OR(H%d="", I%d="", M%d="", AA%d="", AB%d="", AC%d="", AG%d="", AH%d=""), "", ',
          'IF(AG%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(N%d=TRUE, 0, ',
          'IF(I%d<=112, 0, ',
          'IF(M%d<=7, 0, ',
          'IF(H%d<>"Magyar", AC%d*0.15*M%d, AC%d*0.1*M%d) * (M%d/L%d)',
          ')))))',
          ')' # Legkülső IF lezárása
        )
        
        rows <- 13:(13+nrow(konyv)-1)
        write_dynamic_formula(wb, "1_szakkonyv_monog", formula_template, rows, start_col = 32, start_row = 13)

        
        # --- Feltételes formázások
        
        rows_konyv <- 13:(13 + nrow(konyv) - 1)
        apply_cond_format(wb, "1_szakkonyv_monog", cols = 4,  rows = rows_konyv, rule = "D13=INFO!$B$5")
        apply_cond_format(wb, "1_szakkonyv_monog", cols = 8,  rows = rows_konyv, rule = "LEN(H13)=0")
        apply_cond_format(wb, "1_szakkonyv_monog", cols = 9,  rows = rows_konyv, rule = "OR(I13<=112, LEN(I13)=0)")
        apply_cond_format(wb, "1_szakkonyv_monog", cols = 11, rows = rows_konyv, rule = 'K13<>"Tudományos"', bg_color = "#d3d3d3")
        apply_cond_format(wb, "1_szakkonyv_monog", cols = 12, rows = rows_konyv, rule = "LEN(L13)=0")
        apply_cond_format(wb, "1_szakkonyv_monog", cols = 14, rows = rows_konyv, rule = "N13=TRUE")
        apply_cond_format(wb, "1_szakkonyv_monog", cols = 16, rows = rows_konyv, rule = "AND(D13=(INFO!$B$5), P13=FALSE)")
        apply_cond_format(wb, "1_szakkonyv_monog", cols = 17, rows = rows_konyv, rule = 'AND(OR(LEN(Q13)=0, Q13<=7), AI13="Saját ívek száma alapján")')
        apply_cond_format(wb, "1_szakkonyv_monog", cols = 18, rows = rows_konyv, rule = "LEN(R13)=0")
        apply_cond_format(wb, "1_szakkonyv_monog", cols = 23, rows = rows_konyv, rule = "AND(LEN(T13)<>0, LEN(W13)=0)")
        apply_cond_format(wb, "1_szakkonyv_monog", cols = 24, rows = rows_konyv, rule = "AND(LEN(U13)<>0, LEN(X13)=0)")
        apply_cond_format(wb, "1_szakkonyv_monog", cols = 25, rows = rows_konyv, rule = "AND(LEN(V13)<>0, LEN(Y13)=0)")
        apply_cond_format(wb, "1_szakkonyv_monog", cols = 26, rows = rows_konyv, rule = 'AND(LEN($Z13)<>0, $Z13=FALSE)', bg_color = "#DDEBF7")
        apply_cond_format(wb, "1_szakkonyv_monog", cols = 30, rows = rows_konyv, rule = "AND(LEN(AA13)<>0, LEN(AD13)=0)")
        apply_cond_format(wb, "1_szakkonyv_monog", cols = 31, rows = rows_konyv, rule = "AND(LEN(AB13)<>0, LEN(AE13)=0)")
        apply_cond_format(wb, "1_szakkonyv_monog", cols = 32, rows = rows_konyv, rule = "AND(LEN(AC13)<>0, LEN(AF13)=0)")
        apply_cond_format(wb, "1_szakkonyv_monog", cols = 34, rows = rows_konyv, rule = "LEN(AH13)=0")
      
        
        # --- Legördülő listák beállítása
        
        # JELENLEGI hosszú if blokkok helyett csak ennyi:
        add_dropdown(wb, "1_szakkonyv_monog", cols = 14, rows = rows_konyv, value = "'Lists'!$A$1:$A$2")
        add_dropdown(wb, "1_szakkonyv_monog", cols = 16, rows = rows_konyv, value = "'Lists'!$A$1:$A$2")
        add_dropdown(wb, "1_szakkonyv_monog", cols = 33, rows = rows_konyv, value = '"Közlemény és hivatkozások után járó pontszám beszámítása,Csak a hivatkozás(ok) után járó pontszám beszámítása,Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám"')
        add_dropdown(wb, "1_szakkonyv_monog", cols = 35, rows = rows_konyv, value = '"Saját ívek száma alapján,Szerzők számával normált ívszám"')

        
        
        # ---------------------------------------------------------
        # 6. Output tábla feltöltése
        # (2) Folyóiratban megjelent pontozott közlemények
        # ---------------------------------------------------------
        
        log_info("[{session_id}] (2) Folyóiratban megjelent pontozott közlemények")
        
        # --- Közlemények táblázat
        
        # ertekelesi_ter_kod oszlop feltöltése
        publications$ertekelesi_ter_kod[which(publications$otype=="JournalArticle")] <- "2"
        table(publications$ertekelesi_ter_kod)
        
        folyoirat <- publications %>%
          subset(otype=="JournalArticle") %>%
          select(label,mtid,otype,publishedYear,firstAuthor,title,subTitle,
                 volume,issue,firstPage,lastPage,lang,pageLength,
                 subType.label,category.label,
                 categ,nemzetkozi,norveg,citationCount,phd_utan_keletkezett_kozl,ertekelesi_ter_kod,authorCount)
        
        if (nrow(folyoirat)==0) {
          folyoirat[1, ] <- NA          # create one empty row
          folyoirat[1, 1] <- "A jelölt nem rendelkezik folyóiratban megjelent pontozott közleménnyel."
        }
        
        folyoirat <- folyoirat %>% select(label:phd_utan_keletkezett_kozl,ertekelesi_ter_kod,authorCount)
        folyoirat <- folyoirat %>% mutate(pontszám=NA, .before = "ertekelesi_ter_kod")
        colnames(folyoirat)
        
        folyoirat <- folyoirat %>%
          mutate(hiv_nemz_list_folyoirat=NA,
                 hiv_nem_nemz_list_folyoirat=NA,
                 hiv_tov_telj_tud_kozlemeny=NA,
                 hiv_nemz_list_folyoirat_pontszám=NA,
                 hiv_nem_nemz_list_folyoirat_pontszám=NA,
                 hiv_tov_telj_tud_kozlemeny_pontszám=NA,
                 citationCount_eq_hiv=NA)
        
        folyoirat <- folyoirat %>%
          rename(
            "Közlemény" = "label",
            "mtid" = "mtid",
            "Típus" = "otype",
            "Publikálás éve" = "publishedYear",
            "Első szerző" = "firstAuthor",
            "Cím" = "title",
            "Alcím" = "subTitle",
            "Évfolyam" = "volume",
            "Szám" = "issue",
            "Első oldal" = "firstPage",
            "Utolsó oldal" = "lastPage",
            "Nyelv" = "lang",
            "Oldalszám" = "pageLength",
            "Altípus" = "subType.label",
            "Kategória" = "category.label",
            "Lap értékelés kategória" = "categ",
            "Nemzetközi" = "nemzetkozi",
            "Norvég lista" = "norveg",
            "Citációk száma" = "citationCount",
            "PhD után keletkezett közlemény" = "phd_utan_keletkezett_kozl",
            "Pontszám" = "pontszám",
            "Értékelési terület kódja" = "ertekelesi_ter_kod",
            "Hivatkozás nemzetközi listás folyóiratban (db)" = "hiv_nemz_list_folyoirat",
            "Hivatkozás nem nemzetközi listás folyóiratban (db)" = "hiv_nem_nemz_list_folyoirat",
            "Hivatkozás további teljes tudományos közleményben (db)" = "hiv_tov_telj_tud_kozlemeny",
            "Hivatkozás nemzetközi listás folyóiratban (pontszám)" = "hiv_nemz_list_folyoirat_pontszám",
            "Hivatkozás nem nemzetközi listás folyóiratban (pontszám)" = "hiv_nem_nemz_list_folyoirat_pontszám",
            "Hivatkozás további teljes tudományos közleményben (pontszám)" = "hiv_tov_telj_tud_kozlemeny_pontszám",
            "Citációk száma egyezik-e a hivatkozások számával" = "citationCount_eq_hiv",
            "Szerzők száma" = "authorCount"
          ) %>%
          mutate("Hivatkozás a phd fokozat megszerzése óta nemzetközi listás folyóiratban (db)" = NA,
                 "Hivatkozás a phd fokozat megszerzése óta nem nemzetközi listás folyóiratban (db)" = NA,
                 "Hivatkozás a phd fokozat megszerzése óta további teljes tudományos közleményben (db)" = NA,
                 "Hivatkozás a phd fokozat megszerzése óta nemzetközi listás folyóiratban (pontszám)" = NA,
                 "Hivatkozás a phd fokozat megszerzése óta nem nemzetközi listás folyóiratban (pontszám)" = NA,
                 "Hivatkozás a phd fokozat megszerzése óta további teljes tudományos közleményben (pontszám)" = NA) %>%
          mutate(Pontbeszámítás = "Közlemény és hivatkozások után járó pontszám beszámítása") %>%
          relocate("Szerzők száma", .after = "Pontbeszámítás")
        
        folyoirat$Pontbeszámítás[which(folyoirat$Kategória!="Tudományos")] <- "Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám"
        
        writeData(wb, "2_folyoirat", x = folyoirat, startCol = 1, startRow = 20)
        
        
        # --- Hiperhivatkozások
        
        add_mtmt_hyperlinks(wb, "2_folyoirat", df = folyoirat, id_col = "mtid", start_row = 21)
        
        
        # --- phd_utan_keletkezett_kozl függvény
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=D%d>(INFO!$B$5)'
        )
        rows <- 21:(21+nrow(folyoirat)-1)
        write_dynamic_formula(wb, "2_folyoirat", formula_template, rows, start_col = 20, start_row = 21)
        
        
        # --- Közlemények után járó pontszám képlet
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(OR(M%d="", P%d="", Q%d="", AJ%d="", AK%d=""), "", ',
          # Az AK oszlop szerinti szorzó meghatározása (alapértelmezett: 1)
          'IF(AK%d=2, 0.6, IF(AK%d=3, 0.4, IF(AK%d>3, 0.3, 1))) * ',
          # A meglévő teljes pontszámítási logika zárójelben
          '(IF(R%d=TRUE, 0, ',
          'IF(AJ%d<>"Közlemény és hivatkozások után járó pontszám beszámítása", 0, ',
          'IF(AND(Q%d=TRUE, M%d>7, P%d="A"), 18, ',
          'IF(AND(Q%d=TRUE, M%d>7, P%d="B"), 13, ',
          'IF(AND(Q%d=TRUE, M%d>7, P%d="C"), 8, ',
          'IF(AND(Q%d=TRUE, M%d>7, P%d="D"), 6, ',
          'IF(AND(Q%d=TRUE, M%d<=7, M%d>3, P%d="A"), 18*0.5, ',
          'IF(AND(Q%d=TRUE, M%d<=7, M%d>3, P%d="B"), 13*0.5, ',
          'IF(AND(Q%d=TRUE, M%d<=7, M%d>3, P%d="C"), 8*0.5, ',
          'IF(AND(Q%d=TRUE, M%d<=7, M%d>3, P%d="D"), 6*0.5, ',
          'IF(AND(Q%d=TRUE, M%d<=3, P%d="A"), 18*0, ',
          'IF(AND(Q%d=TRUE, M%d<=3, P%d="B"), 13*0, ',
          'IF(AND(Q%d=TRUE, M%d<=3, P%d="C"), 8*0, ',
          'IF(AND(Q%d=TRUE, M%d<=3, P%d="D"), 6*0, ',
          'IF(AND(Q%d=FALSE, M%d>7, P%d="A"), 9, ',
          'IF(AND(Q%d=FALSE, M%d>7, P%d="B"), 6, ',
          'IF(AND(Q%d=FALSE, M%d>7, P%d="C"), 4, ',
          'IF(AND(Q%d=FALSE, M%d>7, P%d="D"), 2, ',
          'IF(AND(Q%d=FALSE, M%d<=7, M%d>3, P%d="A"), 9*0.5, ',
          'IF(AND(Q%d=FALSE, M%d<=7, M%d>3, P%d="B"), 6*0.5, ',
          'IF(AND(Q%d=FALSE, M%d<=7, M%d>3, P%d="C"), 4*0.5, ',
          'IF(AND(Q%d=FALSE, M%d<=7, M%d>3, P%d="D"), 2*0.5, ',
          'IF(AND(Q%d=FALSE, M%d<=3, P%d="A"), 9*0, ',
          'IF(AND(Q%d=FALSE, M%d<=3, P%d="B"), 6*0, ',
          'IF(AND(Q%d=FALSE, M%d<=3, P%d="C"), 4*0, ',
          'IF(AND(Q%d=FALSE, M%d<=3, P%d="D"), 2*0, ',
          '""))))))))))))))))))))))))))))'
        )
        
        rows <- 21:(21+nrow(folyoirat)-1)
        write_dynamic_formula(wb, "2_folyoirat", formula_template, rows, start_col = 21, start_row = 21)
        

        # --- Hivatkozások száma - hiv_nemz_list_folyoirat
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "nemzetközi listás folyóiratban", ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 21:(21+nrow(folyoirat)-1)
        write_dynamic_formula(wb, "2_folyoirat", formula_template, rows, start_col = 23, start_row = 21)
        
        
        # --- Hivatkozások száma - hiv_nem_nemz_list_folyoirat
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "nem nemzetközi listás folyóiratban", ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 21:(21+nrow(folyoirat)-1)
        write_dynamic_formula(wb, "2_folyoirat", formula_template, rows, start_col = 24, start_row = 21)
      
        
        # --- Hivatkozások száma - hiv_tov_telj_tud_kozlemeny
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "további teljes tudományos közleményben", ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 21:(21+nrow(folyoirat)-1)
        write_dynamic_formula(wb, "2_folyoirat", formula_template, rows, start_col = 25, start_row = 21)
        
        
        # --- Hivatkozások pontszám: hiv_nemz_list_folyoirat_pontszám
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(OR(M%d="", P%d="", Q%d="", W%d="", AJ%d="", AK%d=""), "", ',
          # AK oszlop szerinti módosító szorzó (Alapértelmezett: 1)
          'IF(AK%d=2, 0.6, IF(AK%d=3, 0.4, IF(AK%d>3, 0.3, 1))) * ',
          # Belső logika kezdete
          '(IF(AJ%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(R%d=TRUE, 0, ',
          # If M <= 3 -> 0
          'IF(M%d<=3, 0, ',
          # Q == TRUE branch
          'IF(OR(Q%d=TRUE, Q%d="TRUE"), ',
          'IF(M%d>7, ',
          'IF(P%d="A", W%d*6, IF(P%d="B", W%d*4, IF(P%d="C", W%d*2, IF(P%d="D", W%d*1.5, "")))), ',
          'IF(AND(M%d<=7, M%d>3), IF(P%d="A", W%d*6*0.5, IF(P%d="B", W%d*4*0.5, IF(P%d="C", W%d*2*0.5, IF(P%d="D", W%d*1.5*0.5, "")))), "" )',
          '), ',
          # Q == FALSE branch
          'IF(OR(Q%d=FALSE, Q%d="FALSE"), ',
          'IF(M%d>7, ',
          'IF(P%d="A", W%d*2, IF(P%d="B", W%d*1.5, IF(P%d="C", W%d*1.25, IF(P%d="D", W%d*1, "")))), ',
          'IF(AND(M%d<=7, M%d>3), IF(P%d="A", W%d*2*0.5, IF(P%d="B", W%d*1.5*0.5, IF(P%d="C", W%d*1.25*0.5, IF(P%d="D", W%d*1*0.5, "")))), "" )',
          '), ',
          '""' , # fallback if Q neither TRUE nor FALSE
          ')))))))'
        )
        
        rows <- 21:(21+nrow(folyoirat)-1)
        write_dynamic_formula(wb, "2_folyoirat", formula_template, rows, start_col = 26, start_row = 21)

        
        # --- Hivatkozások pontszám: hiv_nem_nemz_list_folyoirat_pontszám
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(OR(M%d="", P%d="", Q%d="", X%d="", AJ%d="", AK%d=""), "", ',
          # AK oszlop szerinti módosító szorzó (Alapértelmezett: 1)
          'IF(AK%d=2, 0.6, IF(AK%d=3, 0.4, IF(AK%d>3, 0.3, 1))) * ',
          # Belső logika kezdete
          '(IF(AJ%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(R%d=TRUE, 0, ',
          # If M <= 3 -> 0
          'IF(M%d<=3, 0, ',
          # Q == TRUE branch
          'IF(OR(Q%d=TRUE, Q%d="TRUE"), ',
          'IF(M%d>7, ',
          'IF(P%d="A", X%d*3, IF(P%d="B", X%d*2, IF(P%d="C", X%d*1.5, IF(P%d="D", X%d*1, "")))), ',
          'IF(AND(M%d<=7, M%d>3), IF(P%d="A", X%d*3*0.5, IF(P%d="B", X%d*2*0.5, IF(P%d="C", X%d*1.5*0.5, IF(P%d="D", X%d*1*0.5, "")))), "" )',
          '), ',
          # Q == FALSE branch
          'IF(OR(Q%d=FALSE, Q%d="FALSE"), ',
          'IF(M%d>7, ',
          'IF(P%d="A", X%d*1.5, IF(P%d="B", X%d*1, IF(P%d="C", X%d*0.75, IF(P%d="D", X%d*0.5, "")))), ',
          'IF(AND(M%d<=7, M%d>3), IF(P%d="A", X%d*1.5*0.5, IF(P%d="B", X%d*1*0.5, IF(P%d="C", X%d*0.75*0.5, IF(P%d="D", X%d*0.5*0.5, "")))), "" )',
          '), ',
          '""' , # fallback if Q neither TRUE nor FALSE
          ')))))))'
        )
        
        rows <- 21:(21+nrow(folyoirat)-1)
        write_dynamic_formula(wb, "2_folyoirat", formula_template, rows, start_col = 27, start_row = 21)
        
        
        # --- Hivatkozások pontszám: hiv_tov_telj_tud_kozlemeny_pontszám
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(OR(M%d="", P%d="", Q%d="", Y%d="", AJ%d="", AK%d=""), "", ',
          # AK oszlop szerinti módosító szorzó (Alapértelmezett: 1)
          'IF(AK%d=2, 0.6, IF(AK%d=3, 0.4, IF(AK%d>3, 0.3, 1))) * ',
          # Belső logika kezdete
          '(IF(AJ%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(R%d=TRUE, 0, ',
          # If M <= 3 -> 0
          'IF(M%d<=3, 0, ',
          # Q == TRUE branch
          'IF(OR(Q%d=TRUE, Q%d="TRUE"), ',
          'IF(M%d>7, ',
          'IF(P%d="A", Y%d*2, IF(P%d="B", Y%d*1.5, IF(P%d="C", Y%d*1, IF(P%d="D", Y%d*0.75, "")))), ',
          'IF(AND(M%d<=7, M%d>3), IF(P%d="A", Y%d*2*0.5, IF(P%d="B", Y%d*1.5*0.5, IF(P%d="C", Y%d*1*0.5, IF(P%d="D", Y%d*0.75*0.5, "")))), "" )',
          '), ',
          # Q == FALSE branch
          'IF(OR(Q%d=FALSE, Q%d="FALSE"), ',
          'IF(M%d>7, ',
          'IF(P%d="A", Y%d*1, IF(P%d="B", Y%d*0.75, IF(P%d="C", Y%d*0.5, IF(P%d="D", Y%d*0.25, "")))), ',
          'IF(AND(M%d<=7, M%d>3), IF(P%d="A", Y%d*1*0.5, IF(P%d="B", Y%d*0.75*0.5, IF(P%d="C", Y%d*0.5*0.5, IF(P%d="D", Y%d*0.25*0.5, "")))), "" )',
          '), ',
          '""' , # fallback if Q neither TRUE nor FALSE
          ')))))))'
        )
        
        rows <- 21:(21+nrow(folyoirat)-1)
        write_dynamic_formula(wb, "2_folyoirat", formula_template, rows, start_col = 28, start_row = 21)
      
        
        # --- citationCount_eq_hiv
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=S%d=(W%d+X%d+Y%d)'
        )
        
        rows <- 21:(21+nrow(folyoirat)-1)
        write_dynamic_formula(wb, "2_folyoirat", formula_template, rows, start_col = 29, start_row = 21)
        
        # Feltételes formázás
        conditionalFormatting(
          wb,
          sheet = "2_folyoirat",
          cols = 29,
          rows = 21:(21+nrow(folyoirat)-1),
          type = "expression",
          rule = 'AND(LEN($AC21)<>0, $AC21=FALSE)',
          style = createStyle(fontColour = "black", bgFill = "#DDEBF7")
        )
        
        
        # --- Hivatkozások száma a phd megszerzése óta - hiv_phd_ota_nemz_list_folyoirat
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "nemzetközi listás folyóiratban", ',
          '4_hivatkozasok!$O$25:O10000, ">"&INFO!$B$4, ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 21:(21+nrow(folyoirat)-1)
        write_dynamic_formula(wb, "2_folyoirat", formula_template, rows, start_col = 30, start_row = 21)
        
        
        # --- Hivatkozások száma a phd megszerzése óta - hiv_nem_nemz_list_folyoirat
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "nem nemzetközi listás folyóiratban", ',
          '4_hivatkozasok!$O$25:O10000, ">"&INFO!$B$4, ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 21:(21+nrow(folyoirat)-1)
        write_dynamic_formula(wb, "2_folyoirat", formula_template, rows, start_col = 31, start_row = 21)
        
        
        # --- Hivatkozások száma a phd megszerzése óta - hiv_tov_telj_tud_kozlemeny
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "további teljes tudományos közleményben", ',
          '4_hivatkozasok!$O$25:O10000, ">"&INFO!$B$4, ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 21:(21+nrow(folyoirat)-1)
        write_dynamic_formula(wb, "2_folyoirat", formula_template, rows, start_col = 32, start_row = 21)
        
        
        # --- Hivatkozások pontszám a phd megszerzése óta: hiv_phd_ota_nemz_list_folyoirat
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(OR(M%d="", P%d="", Q%d="", AD%d="", AJ%d="", AK%d=""), "", ',
          # AK oszlop szerinti módosító szorzó (Alapértelmezett: 1)
          'IF(AK%d=2, 0.6, IF(AK%d=3, 0.4, IF(AK%d>3, 0.3, 1))) * ',
          # Belső logika kezdete
          '(IF(AJ%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(R%d=TRUE, 0, ',
          # If M <= 3 -> 0
          'IF(M%d<=3, 0, ',
          # Q == TRUE branch
          'IF(OR(Q%d=TRUE, Q%d="TRUE"), ',
          'IF(M%d>7, ',
          'IF(P%d="A", AD%d*6, IF(P%d="B", AD%d*4, IF(P%d="C", AD%d*2, IF(P%d="D", AD%d*1.5, "")))), ',
          'IF(AND(M%d<=7, M%d>3), IF(P%d="A", AD%d*6*0.5, IF(P%d="B", AD%d*4*0.5, IF(P%d="C", AD%d*2*0.5, IF(P%d="D", AD%d*1.5*0.5, "")))), "" )',
          '), ',
          # Q == FALSE branch
          'IF(OR(Q%d=FALSE, Q%d="FALSE"), ',
          'IF(M%d>7, ',
          'IF(P%d="A", AD%d*2, IF(P%d="B", AD%d*1.5, IF(P%d="C", AD%d*1.25, IF(P%d="D", AD%d*1, "")))), ',
          'IF(AND(M%d<=7, M%d>3), IF(P%d="A", AD%d*2*0.5, IF(P%d="B", AD%d*1.5*0.5, IF(P%d="C", AD%d*1.25*0.5, IF(P%d="D", AD%d*1*0.5, "")))), "" )',
          '), ',
          '""' , # fallback if Q neither TRUE nor FALSE
          ')))))))'
        )
        
        rows <- 21:(21+nrow(folyoirat)-1)
        write_dynamic_formula(wb, "2_folyoirat", formula_template, rows, start_col = 33, start_row = 21)
        
        
        # --- Hivatkozások pontszám a phd megszerzése óta: hiv_nem_nemz_list_folyoirat
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(OR(M%d="", P%d="", Q%d="", AE%d="", AJ%d="", AK%d=""), "", ',
          # AK oszlop szerinti módosító szorzó (Alapértelmezett: 1)
          'IF(AK%d=2, 0.6, IF(AK%d=3, 0.4, IF(AK%d>3, 0.3, 1))) * ',
          # Belső logika kezdete
          '(IF(AJ%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(R%d=TRUE, 0, ',
          # If M <= 3 -> 0
          'IF(M%d<=3, 0, ',
          # Q == TRUE branch
          'IF(OR(Q%d=TRUE, Q%d="TRUE"), ',
          'IF(M%d>7, ',
          'IF(P%d="A", AE%d*3, IF(P%d="B", AE%d*2, IF(P%d="C", AE%d*1.5, IF(P%d="D", AE%d*1, "")))), ',
          'IF(AND(M%d<=7, M%d>3), IF(P%d="A", AE%d*3*0.5, IF(P%d="B", AE%d*2*0.5, IF(P%d="C", AE%d*1.5*0.5, IF(P%d="D", AE%d*1*0.5, "")))), "" )',
          '), ',
          # Q == FALSE branch
          'IF(OR(Q%d=FALSE, Q%d="FALSE"), ',
          'IF(M%d>7, ',
          'IF(P%d="A", AE%d*1.5, IF(P%d="B", AE%d*1, IF(P%d="C", AE%d*0.75, IF(P%d="D", AE%d*0.5, "")))), ',
          'IF(AND(M%d<=7, M%d>3), IF(P%d="A", AE%d*1.5*0.5, IF(P%d="B", AE%d*1*0.5, IF(P%d="C", AE%d*0.75*0.5, IF(P%d="D", AE%d*0.5*0.5, "")))), "" )',
          '), ',
          '""' , # fallback if Q neither TRUE nor FALSE
          ')))))))'
        )
        
        rows <- 21:(21+nrow(folyoirat)-1)
        write_dynamic_formula(wb, "2_folyoirat", formula_template, rows, start_col = 34, start_row = 21)
      
        
        # --- Hivatkozások pontszám a phd megszerzése óta: hiv_tov_telj_tud_kozlemeny
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(OR(M%d="", P%d="", Q%d="", AF%d="", AJ%d="", AK%d=""), "", ',
          # AK oszlop szerinti módosító szorzó (Alapértelmezett: 1)
          'IF(AK%d=2, 0.6, IF(AK%d=3, 0.4, IF(AK%d>3, 0.3, 1))) * ',
          # Belső logika kezdete
          '(IF(AJ%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          'IF(R%d=TRUE, 0, ',
          # If M <= 3 -> 0
          'IF(M%d<=3, 0, ',
          # Q == TRUE branch
          'IF(OR(Q%d=TRUE, Q%d="TRUE"), ',
          'IF(M%d>7, ',
          'IF(P%d="A", AF%d*2, IF(P%d="B", AF%d*1.5, IF(P%d="C", AF%d*1, IF(P%d="D", AF%d*0.75, "")))), ',
          'IF(AND(M%d<=7, M%d>3), IF(P%d="A", AF%d*2*0.5, IF(P%d="B", AF%d*1.5*0.5, IF(P%d="C", AF%d*1*0.5, IF(P%d="D", AF%d*0.75*0.5, "")))), "" )',
          '), ',
          # Q == FALSE branch
          'IF(OR(Q%d=FALSE, Q%d="FALSE"), ',
          'IF(M%d>7, ',
          'IF(P%d="A", AF%d*1, IF(P%d="B", AF%d*0.75, IF(P%d="C", AF%d*0.5, IF(P%d="D", AF%d*0.25, "")))), ',
          'IF(AND(M%d<=7, M%d>3), IF(P%d="A", AF%d*1*0.5, IF(P%d="B", AF%d*0.75*0.5, IF(P%d="C", AF%d*0.5*0.5, IF(P%d="D", AF%d*0.25*0.5, "")))), "" )',
          '), ',
          '""' , # fallback if Q neither TRUE nor FALSE
          ')))))))'
        )
        
        rows <- 21:(21+nrow(folyoirat)-1)
        write_dynamic_formula(wb, "2_folyoirat", formula_template, rows, start_col = 35, start_row = 21)
        
        
        # --- Feltételes formázások
        
        rows_folyoirat <- 21:(21+nrow(folyoirat)-1)

        apply_cond_format(wb, "2_folyoirat", cols = 4,  rows = rows_folyoirat, rule = "D21=INFO!$B$5")
        apply_cond_format(wb, "2_folyoirat", cols = 13, rows = rows_folyoirat, rule = "LEN(M21)=0")
        apply_cond_format(wb, "2_folyoirat", cols = 15, rows = rows_folyoirat, rule = 'O21<>"Tudományos"', bg_color = "#d3d3d3")
        apply_cond_format(wb, "2_folyoirat", cols = 16, rows = rows_folyoirat, rule = "LEN(P21)=0", bg_color = "#fed8b1")
        apply_cond_format(wb, "2_folyoirat", cols = 17, rows = rows_folyoirat, rule = "LEN(Q21)=0")
        apply_cond_format(wb, "2_folyoirat", cols = 18, rows = rows_folyoirat, rule = "R21=TRUE")
        apply_cond_format(wb, "2_folyoirat", cols = 20, rows = rows_folyoirat, rule = "AND(D21=(INFO!$B$5), T21=FALSE)")
        apply_cond_format(wb, "2_folyoirat", cols = 21, rows = rows_folyoirat, rule = "LEN(U21)=0")
        apply_cond_format(wb, "2_folyoirat", cols = 37, rows = rows_folyoirat, rule = "LEN(AK21)=0")
        apply_cond_format(wb, "2_folyoirat", cols = 26:28, rows = rows_folyoirat, rule = "AND(LEN(W21)<>0, LEN(Z21)=0)")
        apply_cond_format(wb, "2_folyoirat", cols = 33, rows = rows_folyoirat, rule = "AND(LEN(AD21)<>0, LEN(AG21)=0)")
        apply_cond_format(wb, "2_folyoirat", cols = 34, rows = rows_folyoirat, rule = "AND(LEN(AE21)<>0, LEN(AH21)=0)")
        apply_cond_format(wb, "2_folyoirat", cols = 35, rows = rows_folyoirat, rule = "AND(LEN(AF21)<>0, LEN(AI21)=0)")
        
        
        # --- Legördülő listák beállítása
        
        add_dropdown(wb, "2_folyoirat", cols = 16, rows = rows_folyoirat, value = '"A,B,C,D"')
        add_dropdown(wb, "2_folyoirat", cols = 17, rows = rows_folyoirat, value = "'Lists'!$A$1:$A$2")
        add_dropdown(wb, "2_folyoirat", cols = 18, rows = rows_folyoirat, value = "'Lists'!$A$1:$A$2")
        add_dropdown(wb, "2_folyoirat", cols = 20, rows = rows_folyoirat, value = "'Lists'!$A$1:$A$2")
        add_dropdown(wb, "2_folyoirat", cols = 36, rows = rows_folyoirat, value = '"Közlemény és hivatkozások után járó pontszám beszámítása,Csak a hivatkozás(ok) után járó pontszám beszámítása,Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám"')
        
        
        # ---------------------------------------------------------
        # 6. Output tábla feltöltése
        # (3) További pontozott közlemények (csak a 3.1-3.4 összegei!)
        # ---------------------------------------------------------
        
        log_info("[{session_id}] (3) További pontozott közlemények")
        
        # --- További pontozott közlemények táblázat elkészítése
        # 3.1. További (nem kiemelten értékelt) szakkönyv, monográfia (ISBN számmal)
        # Ha könyv és a categ nem ismert (=tehát nem kiemelten értékelt)
        
        publications$ertekelesi_ter_kod[which(publications$ertekelesi_ter_kod=="0" &
                                                publications$otype=="Book" &
                                                is.na(publications$categ)==T &
                                                is.na(publications$isbn_extracted)==F)] <- "3.1"
        
        # 3.2 Szakkönyv szerkesztése (ISBN számmal)
        # Nem leprogramozható
        
        # 3.3. Tudományos könyvrész (ISBN számmal)
        publications$ertekelesi_ter_kod[which(publications$ertekelesi_ter_kod=="0" &
                                                publications$otype=="BookChapter" &
                                                is.na(publications$isbn_extracted)==F)] <- "3.3"
        
        # 3.4. Konferenciaközlemény ISBN-es kötetben, ill. ISSN-es folyóiratban
        # Konferenciaközlemény ISBN-es kötetben
        publications$ertekelesi_ter_kod[which(publications$otype=="BookChapter" &
                                                publications$book.conferencePublication==T &
                                                (publications$subType.label=="Konferenciaközlemény (Könyvrészlet)" |
                                                   publications$subType.label=="Absztrakt / Kivonat (Egyéb konferenciaközlemény)") &
                                                is.na(publications$isbn_extracted)==F)] <- "3.4"
        # Konferenciaközlemény ISSN-es folyóiratban
        publications$ertekelesi_ter_kod[which(publications$otype=="BookChapter" &
                                                publications$book.conferencePublication==T &
                                                (publications$subType.label=="Konferenciaközlemény (Könyvrészlet)" |
                                                   publications$subType.label=="Absztrakt / Kivonat (Egyéb konferenciaközlemény)") &
                                                (is.na(publications$journal.pIssn)==F | is.na(publications$journal.eIssn)==F))] <- "3.4"
        
        #x <- publications %>%
        #  subset(ertekelesi_ter_kod=="0") %>%
        #  select(otype,mtid,label,ertekelesi_ter_kod,book.conferencePublication,subType.label,isbn_extracted,journal.pIssn,journal.eIssn)
        
        # 3.5. További nem pontozott tudományos közlemény
        # Minden más:
        publications$ertekelesi_ter_kod[which(publications$ertekelesi_ter_kod=="0")] <- "3.5"
        
        table(publications$ertekelesi_ter_kod)
        
        tovabbi <- publications %>%
          subset(ertekelesi_ter_kod == "0" |
                   ertekelesi_ter_kod == "3.1" |
                   ertekelesi_ter_kod == "3.3" |
                   ertekelesi_ter_kod == "3.4" |
                   ertekelesi_ter_kod == "3.5") %>%
          select(label,mtid,otype,publishedYear,firstAuthor,title,subTitle,lang,pageLength,book.conferencePublication,
                 subType.label,category.label,categ,nemzetkozi,norveg,citationCount,phd_utan_keletkezett_kozl,
                 ertekelesi_ter_kod,authorCount) %>%
          mutate(ív=NA, .before = "ertekelesi_ter_kod") %>%
          mutate(pontszám=NA, .before = "ertekelesi_ter_kod")
        colnames(tovabbi)
        
        if (nrow(tovabbi)==0) {
          tovabbi[1, ] <- NA          # create one empty row
          tovabbi[1, 1] <- "A jelölt nem rendelkezik további pontozott közleménnyel."
        }
        
        tovabbi <- tovabbi %>%
          mutate(hiv_nemz_list_folyoirat=NA,
                 hiv_nem_nemz_list_folyoirat=NA,
                 hiv_tov_telj_tud_kozlemeny=NA,
                 hiv_nemz_list_folyoirat_pontszám=NA,
                 hiv_nem_nemz_list_folyoirat_pontszám=NA,
                 hiv_tov_telj_tud_kozlemeny_pontszám=NA,
                 citationCount_eq_hiv=NA)
        colnames(tovabbi)
        tovabbi <- tovabbi %>%
          rename(
            "Közlemény" = "label",
            "mtid" = "mtid",
            "Típus" = "otype",
            "Publikálás éve" = "publishedYear",
            "Első szerző" = "firstAuthor",
            "Cím" = "title",
            "Alcím" = "subTitle",
            "Nyelv" = "lang",
            "Oldalszám" = "pageLength",
            "Konferencia publikáció könyv" = "book.conferencePublication",
            "Altípus" = "subType.label",
            "Kategória" = "category.label",
            "Lap értékelés kategória" = "categ",
            "Nemzetközi" = "nemzetkozi",
            "Norvég lista" = "norveg",
            "Citációk száma" = "citationCount",
            "PhD után keletkezett közlemény" = "phd_utan_keletkezett_kozl",
            "Saját ívek száma" = "ív",
            "Pontszám" = "pontszám",
            "Értékelési terület kódja" = "ertekelesi_ter_kod",
            "Hivatkozás nemzetközi listás folyóiratban (db)" = "hiv_nemz_list_folyoirat",
            "Hivatkozás nem nemzetközi listás folyóiratban (db)" = "hiv_nem_nemz_list_folyoirat",
            "Hivatkozás további teljes tudományos közleményben (db)" = "hiv_tov_telj_tud_kozlemeny",
            "Hivatkozás nemzetközi listás folyóiratban (pontszám)" = "hiv_nemz_list_folyoirat_pontszám",
            "Hivatkozás nem nemzetközi listás folyóiratban (pontszám)" = "hiv_nem_nemz_list_folyoirat_pontszám",
            "Hivatkozás további teljes tudományos közleményben (pontszám)" = "hiv_tov_telj_tud_kozlemeny_pontszám",
            "Citációk száma egyezik-e a hivatkozások számával" = "citationCount_eq_hiv",
            "Szerzők száma" = "authorCount"
          ) %>%
          mutate("Hivatkozás a phd fokozat megszerzése óta nemzetközi listás folyóiratban (db)" = NA,
                 "Hivatkozás a phd fokozat megszerzése óta nem nemzetközi listás folyóiratban (db)" = NA,
                 "Hivatkozás a phd fokozat megszerzése óta további teljes tudományos közleményben (db)" = NA,
                 "Hivatkozás a phd fokozat megszerzése óta nemzetközi listás folyóiratban (pontszám)" = NA,
                 "Hivatkozás a phd fokozat megszerzése óta nem nemzetközi listás folyóiratban (pontszám)" = NA,
                 "Hivatkozás a phd fokozat megszerzése óta további teljes tudományos közleményben (pontszám)" = NA) %>%
          mutate(Pontbeszámítás = "Közlemény és hivatkozások után járó pontszám beszámítása") %>%
          relocate("Szerzők száma", .after = "Pontbeszámítás")
        
        tovabbi$Pontbeszámítás[which(tovabbi$`Értékelési terület kódja` == "3.1")] <- "Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám"
        
        writeData(wb, "3_tovabbi", x = tovabbi, startCol = 1, startRow = 22)
        
        
        # --- Hiperhivatkozások
        
        add_mtmt_hyperlinks(wb, "3_tovabbi", df = tovabbi, id_col = "mtid", start_row = 23)
        
        
        # --- ertekelesi_ter_kod oszlop szöveggé alakítása
        
        text_style <- createStyle(numFmt = "@")  # "@" means text format in Excel
        addStyle(
          wb,
          sheet = "3_tovabbi",
          style = text_style,
          rows = 23:100,
          cols = 20,       # Column T = 20
          gridExpand = TRUE,
          stack = TRUE
        )
        
        
        # --- phd_utan_keletkezett_kozl függvény
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=D%d>(INFO!$B$5)'
        )
        rows <- 23:(23+nrow(tovabbi)-1)
        write_dynamic_formula(wb, "3_tovabbi", formula_template, rows, start_col = 17, start_row = 23)

        
        # --- Közlemények után járó pontszám képlet
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(AND(T%d="3.4", OR(H%d="", I%d="", T%d="", AH%d="", AI%d="")), "", ',
          'IF(AND(T%d="3.3", I%d<=7), 0, ',
          'IF(AND(T%d="3.3", OR(H%d="", I%d="", R%d="", T%d="", AH%d="", AI%d="")), "", ',
          'IF(AND(OR(T%d="3.1", T%d="3.2"), OR(R%d="", AH%d="", AI%d="")), "", ',
          
          # AI oszlop szerinti módosító szorzó bevezetése (Alapértelmezett: 1)
          'IF(AI%d=2, 0.6, IF(AI%d=3, 0.4, IF(AI%d>3, 0.3, 1))) * ',
          
          # Belső pontszámítási blokk kezdete
          '(IF(O%d=TRUE, 0, ',
          'IF(AH%d<>"Közlemény és hivatkozások után járó pontszám beszámítása", 0, ',
          
          # 3.1
          'IF(AND(T%d="3.1", R%d<=3), 0, ',
          'IF(AND(T%d="3.1", H%d<>"Magyar", R%d>3), R%d*2, ',
          'IF(AND(T%d="3.1", H%d="Magyar", R%d>3), R%d, ',
          
          # 3.2
          'IF(AND(T%d="3.2", R%d<=7), 0, ',
          'IF(AND(T%d="3.2", H%d<>"Magyar", R%d>7), 3, ',
          'IF(AND(T%d="3.2", H%d="Magyar", R%d>7), 2, ',
          
          # 3.3
          'IF(AND(T%d="3.3", H%d<>"Magyar", I%d>7, I%d<=16), R%d*2*0.5, ',
          'IF(AND(T%d="3.3", H%d="Magyar", I%d>7, I%d<=16), R%d*1*0.5, ',
          'IF(AND(T%d="3.3", H%d<>"Magyar", I%d>16), R%d*2, ',
          'IF(AND(T%d="3.3", H%d="Magyar", I%d>16), R%d*1, ',
          
          # 3.4
          'IF(AND(T%d="3.4", I%d<=7), 0, ',
          'IF(AND(T%d="3.4", H%d<>"Magyar", I%d>7), 2, ',
          'IF(AND(T%d="3.4", H%d="Magyar", I%d>7), 1, ',
          
          # 3.5
          'IF(T%d="3.5", 0, ',
          
          '"")))))))))))))))))))))'
        )
        
        rows <- 23:(23+nrow(tovabbi)-1)
        write_dynamic_formula(wb, "3_tovabbi", formula_template, rows, start_col = 19, start_row = 23)
      
        
        # --- Hivatkozások száma - hiv_nemz_list_folyoirat
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "nemzetközi listás folyóiratban", ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 23:(23+nrow(tovabbi)-1)
        write_dynamic_formula(wb, "3_tovabbi", formula_template, rows, start_col = 21, start_row = 23)
        
        
        # --- Hivatkozások száma - hiv_nem_nemz_list_folyoirat
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "nem nemzetközi listás folyóiratban", ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 23:(23+nrow(tovabbi)-1)
        write_dynamic_formula(wb, "3_tovabbi", formula_template, rows, start_col = 22, start_row = 23)
        
        
        # --- Hivatkozások száma - hiv_tov_telj_tud_kozlemeny
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "további teljes tudományos közleményben", ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 23:(23+nrow(tovabbi)-1)
        write_dynamic_formula(wb, "3_tovabbi", formula_template, rows, start_col = 23, start_row = 23)
        
        
        # --- Hivatkozások pontszám: hiv_nemz_list_folyoirat_pontszám
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(OR(H%d="", T%d="", U%d="", AH%d="", AI%d=""), "", ',
          'IF(AND(T%d="3.1", R%d=""), "", ',
          'IF(AND(T%d="3.3", R%d=""), "", ',
          
          # AI oszlop szerinti módosító szorzó (Alapértelmezett: 1)
          'IF(AI%d=2, 0.6, IF(AI%d=3, 0.4, IF(AI%d>3, 0.3, 1))) * ',
          
          # Pontszámítási blokk kezdete
          '(IF(O%d=TRUE, 0, ',
          'IF(AH%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          
          # 3.1
          'IF(AND(T%d="3.1", R%d<=3), 0, ',
          'IF(AND(T%d="3.1", H%d<>"Magyar", R%d>3), ', 'R%d*U%d*0.4, ',
          'IF(AND(T%d="3.1", H%d="Magyar", R%d>3), ', 'R%d*U%d*0.2, ',
          
          # 3.2
          'IF(T%d="3.2", ', '0, ',
          
          # 3.3
          'IF(AND(T%d="3.3", I%d<=7), 0, ',
          'IF(AND(T%d="3.3", H%d<>"Magyar", I%d>7, I%d<=16), ', 'R%d*U%d*0.4*0.5, ',
          'IF(AND(T%d="3.3", H%d="Magyar", I%d>7, I%d<=16), ', 'R%d*U%d*0.2*0.5, ',
          'IF(AND(T%d="3.3", H%d<>"Magyar", I%d>16), ', 'R%d*U%d*0.4, ',
          'IF(AND(T%d="3.3", H%d="Magyar", I%d>16), ', 'R%d*U%d*0.2, ',
          
          # 3.4
          'IF(AND(T%d="3.4", I%d<=7), 0, ',
          'IF(AND(T%d="3.4", H%d<>"Magyar", I%d>7), ', 'U%d*0.4, ',
          'IF(AND(T%d="3.4", H%d="Magyar", I%d>7), ', 'U%d*0.2, ',
          
          # 3.5
          'IF(T%d="3.5", ', '0, ',
          
          '"")))))))))))))))))))'
        )
        
        rows <- 23:(23+nrow(tovabbi)-1)
        write_dynamic_formula(wb, "3_tovabbi", formula_template, rows, start_col = 24, start_row = 23)
        
       
        # --- Hivatkozások pontszám: hiv_nem_nemz_list_folyoirat_pontszám
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(OR(H%d="", T%d="", V%d="", AH%d="", AI%d=""), "", ',
          'IF(AND(T%d="3.1", R%d=""), "", ',
          'IF(AND(T%d="3.3", R%d=""), "", ',
          
          # AI oszlop szerinti módosító szorzó (Alapértelmezett: 1)
          'IF(AI%d=2, 0.6, IF(AI%d=3, 0.4, IF(AI%d>3, 0.3, 1))) * ',
          
          # Pontszámítási blokk kezdete
          '(IF(O%d=TRUE, 0, ',
          'IF(AH%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          
          # 3.1
          'IF(AND(T%d="3.1", R%d<=3), 0, ',
          'IF(AND(T%d="3.1", H%d<>"Magyar", R%d>3), ', 'R%d*V%d*0.2, ',
          'IF(AND(T%d="3.1", H%d="Magyar", R%d>3), ', 'R%d*V%d*0.1, ',
          
          # 3.2
          'IF(T%d="3.2", ', '0, ',
          
          # 3.3
          'IF(AND(T%d="3.3", I%d<=7), 0, ',
          'IF(AND(T%d="3.3", H%d<>"Magyar", I%d>7, I%d<=16), ', 'R%d*V%d*0.2*0.5, ',
          'IF(AND(T%d="3.3", H%d="Magyar", I%d>7, I%d<=16), ', 'R%d*V%d*0.1*0.5, ',
          'IF(AND(T%d="3.3", H%d<>"Magyar", I%d>16), ', 'R%d*V%d*0.2, ',
          'IF(AND(T%d="3.3", H%d="Magyar", I%d>16), ', 'R%d*V%d*0.1, ',
          
          # 3.4:
          'IF(AND(T%d="3.4", I%d<=7), 0, ',
          'IF(AND(T%d="3.4", H%d<>"Magyar", I%d>7), ', 'V%d*0.2, ',
          'IF(AND(T%d="3.4", H%d="Magyar", I%d>7), ', 'V%d*0.1, ',
          
          # 3.5
          'IF(T%d="3.5", ', '0, ',
          
          '"")))))))))))))))))))'
        )
        
        rows <- 23:(23+nrow(tovabbi)-1)
        write_dynamic_formula(wb, "3_tovabbi", formula_template, rows, start_col = 25, start_row = 23)
        

        # --- Hivatkozások pontszám: hiv_tov_telj_tud_kozlemeny_pontszám
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(OR(H%d="", T%d="", W%d="", AH%d="", AI%d=""), "", ',
          'IF(AND(T%d="3.1", R%d=""), "", ',
          'IF(AND(T%d="3.3", R%d=""), "", ',
          
          # AI oszlop szerinti módosító szorzó (Alapértelmezett: 1)
          'IF(AI%d=2, 0.6, IF(AI%d=3, 0.4, IF(AI%d>3, 0.3, 1))) * ',
          
          # Pontszámítási blokk kezdete
          '(IF(O%d=TRUE, 0, ',
          'IF(AH%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          
          # 3.1
          'IF(AND(T%d="3.1", R%d<=3), 0, ',
          'IF(AND(T%d="3.1", H%d<>"Magyar", R%d>3), ', 'R%d*W%d*0.1, ',
          'IF(AND(T%d="3.1", H%d="Magyar", R%d>3), ', 'R%d*W%d*0.05, ',
          
          # 3.2
          'IF(T%d="3.2", ', '0, ',
          
          # 3.3
          'IF(AND(T%d="3.3", I%d<=7), 0, ',
          'IF(AND(T%d="3.3", H%d<>"Magyar", I%d>7, I%d<=16), ', 'R%d*W%d*0.1*0.5, ',
          'IF(AND(T%d="3.3", H%d="Magyar", I%d>7, I%d<=16), ', 'R%d*W%d*0.05*0.5, ',
          'IF(AND(T%d="3.3", H%d<>"Magyar", I%d>16), ', 'R%d*W%d*0.1, ',
          'IF(AND(T%d="3.3", H%d="Magyar", I%d>16), ', 'R%d*W%d*0.05, ',
          
          # 3.4:
          'IF(AND(T%d="3.4", I%d<=7), 0, ',
          'IF(AND(T%d="3.4", H%d<>"Magyar", I%d>7), ', 'W%d*0.1, ',
          'IF(AND(T%d="3.4", H%d="Magyar", I%d>7), ', 'W%d*0.05, ',
          
          # 3.5
          'IF(T%d="3.5", ', '0, ',
          
          '"")))))))))))))))))))'
        )
        
        rows <- 23:(23+nrow(tovabbi)-1)
        write_dynamic_formula(wb, "3_tovabbi", formula_template, rows, start_col = 26, start_row = 23)

        
        # --- citationCount_eq_hiv
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=P%d=(U%d+V%d+W%d)'
        )
        
        rows <- 23:(23+nrow(tovabbi)-1)
        write_dynamic_formula(wb, "3_tovabbi", formula_template, rows, start_col = 27, start_row = 23)

        
        # --- Hivatkozások száma a phd megszerzése óta - hiv_phd_ota_nemz_list_folyoirat
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "nemzetközi listás folyóiratban", ',
          '4_hivatkozasok!$O$25:O10000, ">"&INFO!$B$4, ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 23:(23+nrow(tovabbi)-1)
        write_dynamic_formula(wb, "3_tovabbi", formula_template, rows, start_col = 28, start_row = 23)
        
        
        # --- Hivatkozások száma a phd megszerzése óta - hiv_nem_nemz_list_folyoirat
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "nem nemzetközi listás folyóiratban", ',
          '4_hivatkozasok!$O$25:O10000, ">"&INFO!$B$4, ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 23:(23+nrow(tovabbi)-1)
        write_dynamic_formula(wb, "3_tovabbi", formula_template, rows, start_col = 29, start_row = 23)
        
        
        # --- Hivatkozások száma a phd megszerzése óta - hiv_tov_telj_tud_kozlemeny
        
        formula_template <- paste0(
          '=SUM(COUNTIFS(',
          '4_hivatkozasok!$B$25:B10000, B%d, ',
          '4_hivatkozasok!$I$25:I10000, 0, ',
          '4_hivatkozasok!$N$25:N10000, 0, ',
          '4_hivatkozasok!$J$25:J10000, "további teljes tudományos közleményben", ',
          '4_hivatkozasok!$O$25:O10000, ">"&INFO!$B$4, ',
          '4_hivatkozasok!$E$25:E10000, {"JournalArticle", "Book", "BookChapter"}',
          '))'
        )
        
        rows <- 23:(23+nrow(tovabbi)-1)
        write_dynamic_formula(wb, "3_tovabbi", formula_template, rows, start_col = 30, start_row = 23)
        
        
        # --- Hivatkozások pontszám a phd megszerzése óta: hiv_nemz_list_folyoirat_pontszám
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(OR(H%d="", T%d="", AB%d="", AH%d="", AI%d=""), "", ',
          'IF(AND(T%d="3.1", R%d=""), "", ',
          'IF(AND(T%d="3.3", R%d=""), "", ',
          
          # AI oszlop szerinti módosító szorzó (Alapértelmezett: 1)
          'IF(AI%d=2, 0.6, IF(AI%d=3, 0.4, IF(AI%d>3, 0.3, 1))) * ',
          
          # Pontszámítási blokk kezdete
          '(IF(O%d=TRUE, 0, ',
          'IF(AH%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          
          # 3.1
          'IF(AND(T%d="3.1", R%d<=3), 0, ',
          'IF(AND(T%d="3.1", H%d<>"Magyar", R%d>3), ', 'R%d*U%d*0.4, ',
          'IF(AND(T%d="3.1", H%d="Magyar", R%d>3), ', 'R%d*U%d*0.2, ',
          
          # 3.2
          'IF(T%d="3.2", ', '0, ',
          
          # 3.3
          'IF(AND(T%d="3.3", I%d<=7), 0, ',
          'IF(AND(T%d="3.3", H%d<>"Magyar", I%d>7, I%d<=16), ', 'R%d*U%d*0.4*0.5, ',
          'IF(AND(T%d="3.3", H%d="Magyar", I%d>7, I%d<=16), ', 'R%d*U%d*0.2*0.5, ',
          'IF(AND(T%d="3.3", H%d<>"Magyar", I%d>16), ', 'R%d*U%d*0.4, ',
          'IF(AND(T%d="3.3", H%d="Magyar", I%d>16), ', 'R%d*U%d*0.2, ',
          
          # 3.4:
          'IF(AND(T%d="3.4", I%d<=7), 0, ',
          'IF(AND(T%d="3.4", H%d<>"Magyar", I%d>7), ', 'U%d*0.4, ',
          'IF(AND(T%d="3.4", H%d="Magyar", I%d>7), ', 'U%d*0.2, ',
          
          # 3.5
          'IF(T%d="3.5", ', '0, ',
          
          '"")))))))))))))))))))'
        )
        
        rows <- 23:(23+nrow(tovabbi)-1)
        write_dynamic_formula(wb, "3_tovabbi", formula_template, rows, start_col = 31, start_row = 23)
        
        
        # --- Hivatkozások pontszám a phd megszerzése óta: hiv_nem_nemz_list_folyoirat_pontszám
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(OR(H%d="", T%d="", AC%d="", AH%d="", AI%d=""), "", ',
          'IF(AND(T%d="3.1", R%d=""), "", ',
          'IF(AND(T%d="3.3", R%d=""), "", ',
          
          # AI oszlop szerinti módosító szorzó (Alapértelmezett: 1)
          'IF(AI%d=2, 0.6, IF(AI%d=3, 0.4, IF(AI%d>3, 0.3, 1))) * ',
          
          # Pontszámítási blokk kezdete
          '(IF(O%d=TRUE, 0, ',
          'IF(AH%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          
          # 3.1
          'IF(AND(T%d="3.1", R%d<=3), 0, ',
          'IF(AND(T%d="3.1", H%d<>"Magyar", R%d>3), ', 'R%d*AC%d*0.2, ',
          'IF(AND(T%d="3.1", H%d="Magyar", R%d>3), ', 'R%d*AC%d*0.1, ',
          
          # 3.2
          'IF(T%d="3.2", ', '0, ',
          
          # 3.3
          'IF(AND(T%d="3.3", I%d<=7), 0, ',
          'IF(AND(T%d="3.3", H%d<>"Magyar", I%d>7, I%d<=16), ', 'R%d*AC%d*0.2*0.5, ',
          'IF(AND(T%d="3.3", H%d="Magyar", I%d>7, I%d<=16), ', 'R%d*AC%d*0.1*0.5, ',
          'IF(AND(T%d="3.3", H%d<>"Magyar", I%d>16), ', 'R%d*AC%d*0.2, ',
          'IF(AND(T%d="3.3", H%d="Magyar", I%d>16), ', 'R%d*AC%d*0.1, ',
          
          # 3.4:
          'IF(AND(T%d="3.4", I%d<=7), 0, ',
          'IF(AND(T%d="3.4", H%d<>"Magyar", I%d>7), ', 'AC%d*0.2, ',
          'IF(AND(T%d="3.4", H%d="Magyar", I%d>7), ', 'AC%d*0.1, ',
          
          # 3.5
          'IF(T%d="3.5", ', '0, ',
          
          '"")))))))))))))))))))'
        )
        
        rows <- 23:(23+nrow(tovabbi)-1)
        write_dynamic_formula(wb, "3_tovabbi", formula_template, rows, start_col = 32, start_row = 23)
        
        
        # --- Hivatkozások pontszám a phd megszerzése óta: hiv_tov_telj_tud_kozlemeny_pontszám
        
        # Függvény dinamikus elkészítése az összes sorra
        formula_template <- paste0(
          '=IF(OR(H%d="", T%d="", AD%d="", AH%d="", AI%d=""), "", ',
          'IF(AND(T%d="3.1", R%d=""), "", ',
          'IF(AND(T%d="3.3", R%d=""), "", ',
          
          # AI oszlop szerinti módosító szorzó (Alapértelmezett: 1)
          'IF(AI%d=2, 0.6, IF(AI%d=3, 0.4, IF(AI%d>3, 0.3, 1))) * ',
          
          # Pontszámítási blokk kezdete
          '(IF(O%d=TRUE, 0, ',
          'IF(AH%d="Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám", 0, ',
          
          # 3.1
          'IF(AND(T%d="3.1", R%d<=3), 0, ',
          'IF(AND(T%d="3.1", H%d<>"Magyar", R%d>3), ', 'R%d*AD%d*0.1, ',
          'IF(AND(T%d="3.1", H%d="Magyar", R%d>3), ', 'R%d*AD%d*0.05, ',
          
          # 3.2
          'IF(T%d="3.2", ', '0, ',
          
          # 3.3
          'IF(AND(T%d="3.3", I%d<=7), 0, ',
          'IF(AND(T%d="3.3", H%d<>"Magyar", I%d>7, I%d<=16), ', 'R%d*AD%d*0.1*0.5, ',
          'IF(AND(T%d="3.3", H%d="Magyar", I%d>7, I%d<=16), ', 'R%d*AD%d*0.05*0.5, ',
          'IF(AND(T%d="3.3", H%d<>"Magyar", I%d>16), ', 'R%d*AD%d*0.1, ',
          'IF(AND(T%d="3.3", H%d="Magyar", I%d>16), ', 'R%d*AD%d*0.05, ',
          
          # 3.4:
          'IF(AND(T%d="3.4", I%d<=7), 0, ',
          'IF(AND(T%d="3.4", H%d<>"Magyar", I%d>7), ', 'AD%d*0.1, ',
          'IF(AND(T%d="3.4", H%d="Magyar", I%d>7), ', 'AD%d*0.05, ',
          
          # 3.5
          'IF(T%d="3.5", ', '0, ',
          
          '"")))))))))))))))))))'
        )
        
        rows <- 23:(23+nrow(tovabbi)-1)
        write_dynamic_formula(wb, "3_tovabbi", formula_template, rows, start_col = 33, start_row = 23)
        
        
        # --- Feltételes formázások
        
        rows_tovabbi <- 23:(23+nrow(tovabbi)-1)
        apply_cond_format(wb, "3_tovabbi", cols = 4,  rows = rows_tovabbi, rule = "D23=INFO!$B$5")
        apply_cond_format(wb, "3_tovabbi", cols = 8,  rows = rows_tovabbi, rule = 'AND(LEN(H23)=0, OR(T23="3.1", T23="3.2", T23="3.3", T23="3.4"))')
        apply_cond_format(wb, "3_tovabbi", cols = 9,  rows = rows_tovabbi, rule = 'AND(LEN(I23)=0, OR(T23="3.3", T23="3.4"))')
        apply_cond_format(wb, "3_tovabbi", cols = 12, rows = rows_tovabbi, rule = 'L23<>"Tudományos"', bg_color = "#d3d3d3")
        apply_cond_format(wb, "3_tovabbi", cols = 15, rows = rows_tovabbi, rule = "O23=TRUE")
        apply_cond_format(wb, "3_tovabbi", cols = 17, rows = rows_tovabbi, rule = "AND(D23=(INFO!$B$5), Q23=FALSE)")
        apply_cond_format(wb, "3_tovabbi", cols = 18, rows = rows_tovabbi, rule = 'OR(AND(R23<=3, T23="3.1"), AND(R23<=7, T23="3.2"), AND(LEN(S23)=0, LEN(R23)=0, OR(T23="3.1", T23="3.2", T23="3.3")))')
        apply_cond_format(wb, "3_tovabbi", cols = 19, rows = rows_tovabbi, rule = 'AND(LEN(S23)=0, T23<>"3.5")')
        apply_cond_format(wb, "3_tovabbi", cols = 20, rows = rows_tovabbi, rule = 'AND(LEN(S23)=0, LEN(T23)=0)')
        apply_cond_format(wb, "3_tovabbi", cols = 24, rows = rows_tovabbi, rule = "AND(LEN(U23)<>0, LEN(X23)=0)")
        apply_cond_format(wb, "3_tovabbi", cols = 25, rows = rows_tovabbi, rule = "AND(LEN(V23)<>0, LEN(Y23)=0)")
        apply_cond_format(wb, "3_tovabbi", cols = 26, rows = rows_tovabbi, rule = "AND(LEN(W23)<>0, LEN(Z23)=0)")
        apply_cond_format(wb, "3_tovabbi", cols = 27, rows = rows_tovabbi, rule = 'AND(LEN($AA23)<>0, $AA23=FALSE)', bg_color = "#DDEBF7")
        apply_cond_format(wb, "3_tovabbi", cols = 31, rows = rows_tovabbi, rule = "AND(LEN(AB23)<>0, LEN(AE23)=0)")
        apply_cond_format(wb, "3_tovabbi", cols = 32, rows = rows_tovabbi, rule = "AND(LEN(AC23)<>0, LEN(AF23)=0)")
        apply_cond_format(wb, "3_tovabbi", cols = 33, rows = rows_tovabbi, rule = "AND(LEN(AD23)<>0, LEN(AG23)=0)")
        apply_cond_format(wb, "3_tovabbi", cols = 35, rows = rows_tovabbi, rule = "LEN(AI23)=0")
        
        
        # --- Legördülő listák beállítása
        
        add_dropdown(wb, "3_tovabbi", cols = 13, rows = rows_tovabbi, value = '"A,B,C,D"')
        add_dropdown(wb, "3_tovabbi", cols = 14, rows = rows_tovabbi, value = "'Lists'!$A$1:$A$2")
        add_dropdown(wb, "3_tovabbi", cols = 15, rows = rows_tovabbi, value = "'Lists'!$A$1:$A$2")
        add_dropdown(wb, "3_tovabbi", cols = 17, rows = rows_tovabbi, value = "'Lists'!$A$1:$A$2")
        add_dropdown(wb, "3_tovabbi", cols = 20, rows = rows_tovabbi, value = '"3.1,3.2,3.3,3.4,3.5"')
        add_dropdown(wb, "3_tovabbi", cols = 34, rows = rows_tovabbi, value = '"Közlemény és hivatkozások után járó pontszám beszámítása,Csak a hivatkozás(ok) után járó pontszám beszámítása,Nem kerül beszámításra se a publikációért se a hivatkozás(ok)ért járó pontszám"')


        
        # ---------------------------------------------------------
        # 7. Munkafüzet védelme és cellák zárolása
        # ---------------------------------------------------------
        
        stilus_nyitott <- createStyle(locked = FALSE)
        stilus_zart    <- createStyle(locked = TRUE)
        
        # --- INFO
        
        # Minden cella lezárása
        max_sor <- 200
        max_oszlop <- 2
        
        addStyle(
          wb = wb,
          sheet = "INFO",
          style = stilus_zart,
          rows = 1:max_sor,
          cols = 1:max_oszlop,
          gridExpand = TRUE,
          stack = TRUE              # MEGŐRZI a színeket/formázást!
        )
        
        # A kívánt oszlop nyitása
        addStyle(
          wb = wb,
          sheet = "INFO",
          style = stilus_nyitott,
          rows = 4,
          cols = 2,
          gridExpand = TRUE,
          stack = TRUE
        )
        
        protectWorksheet(wb = wb, sheet = "INFO", protect = TRUE)
        
        
        # --- OSSZEGZES
        
        # Minden cella lezárása
        max_sor <- 200
        max_oszlop <- 5
        
        addStyle(
          wb = wb,
          sheet = "OSSZEGZES",
          style = stilus_zart,
          rows = 1:max_sor,
          cols = 1:max_oszlop,
          gridExpand = TRUE,
          stack = TRUE              # MEGŐRZI a színeket/formázást!
        )
        
        protectWorksheet(wb = wb, sheet = "OSSZEGZES", protect = TRUE)
        
        
        # --- 1_szakkonyv_monog
        
        # Minden cella feloldása a munkaterületen
        max_sor <- (13+nrow(konyv)-1) + 100  # Ráhagyunk 100 sort
        max_oszlop <- ncol(konyv)
        
        addStyle(
          wb = wb,
          sheet = "1_szakkonyv_monog",
          style = stilus_nyitott,
          rows = 1:max_sor,
          cols = 1:max_oszlop,
          gridExpand = TRUE,
          stack = TRUE              # MEGŐRZI a színeket/formázást!
        )
        
        # A kívánt oszlop visszazárása
        addStyle(
          wb = wb,
          sheet = "1_szakkonyv_monog",
          style = stilus_zart,
          rows = 13:max_sor,
          cols = c(1:7, 18, 20:25, 27:32),
          gridExpand = TRUE,
          stack = TRUE
        )
        
        # A kívánt oszlop visszazárása
        addStyle(
          wb = wb,
          sheet = "1_szakkonyv_monog",
          style = stilus_zart,
          rows = 1:8,
          cols = 1:5,
          gridExpand = TRUE,
          stack = TRUE
        )
        
        # Lapvédelem aktiválása
        protectWorksheet(wb = wb, sheet = "1_szakkonyv_monog", protect = TRUE)
        
        
        # --- 2_folyoirat
        
        # Minden cella feloldása a munkaterületen
        max_sor <- (21+nrow(folyoirat)-1) + 100  # Ráhagyunk 100 sort
        max_oszlop <- ncol(folyoirat)
        
        addStyle(
          wb = wb,
          sheet = "2_folyoirat",
          style = stilus_nyitott,
          rows = 1:max_sor,
          cols = 1:max_oszlop,
          gridExpand = TRUE,
          stack = TRUE              # MEGŐRZI a színeket/formázást!
        )
        
        # A kívánt oszlop visszazárása
        addStyle(
          wb = wb,
          sheet = "2_folyoirat",
          style = stilus_zart,
          rows = 21:max_sor,
          cols = c(1:12, 21, 23:28, 30:35),
          gridExpand = TRUE,
          stack = TRUE
        )
        
        # A kívánt oszlop visszazárása
        addStyle(
          wb = wb,
          sheet = "2_folyoirat",
          style = stilus_zart,
          rows = 1:16,
          cols = 1:5,
          gridExpand = TRUE,
          stack = TRUE
        )
        
        # Lapvédelem aktiválása
        protectWorksheet(wb = wb, sheet = "2_folyoirat", protect = TRUE)
        
        
        # --- 3_tovabbi
        
        # Minden cella feloldása a munkaterületen
        max_sor <- (23+nrow(tovabbi)-1) + 100  # Ráhagyunk 100 sort
        max_oszlop <- ncol(tovabbi)
        
        addStyle(
          wb = wb,
          sheet = "3_tovabbi",
          style = stilus_nyitott,
          rows = 1:max_sor,
          cols = 1:max_oszlop,
          gridExpand = TRUE,
          stack = TRUE              # MEGŐRZI a színeket/formázást!
        )
        
        # A kívánt oszlop visszazárása
        addStyle(
          wb = wb,
          sheet = "3_tovabbi",
          style = stilus_zart,
          rows = 23:max_sor,
          cols = c(1:7, 19, 21:26, 28:33),
          gridExpand = TRUE,
          stack = TRUE
        )
        
        # A kívánt oszlop visszazárása
        addStyle(
          wb = wb,
          sheet = "3_tovabbi",
          style = stilus_zart,
          rows = 1:18,
          cols = 1:5,
          gridExpand = TRUE,
          stack = TRUE
        )
        
        # Lapvédelem aktiválása
        protectWorksheet(wb = wb, sheet = "3_tovabbi", protect = TRUE)
        
        
        # --- 4_hivatkozasok
        
        # Minden cella feloldása a munkaterületen
        max_sor <- (26+nrow(all_citations)-1) + 100  # Ráhagyunk 100 sort
        max_oszlop <- ncol(all_citations)
        
        addStyle(
          wb = wb,
          sheet = "4_hivatkozasok",
          style = stilus_nyitott,
          rows = 1:max_sor,
          cols = 1:max_oszlop,
          gridExpand = TRUE,
          stack = TRUE              # MEGŐRZI a színeket/formázást!
        )
        
        # A kívánt oszlop visszazárása
        addStyle(
          wb = wb,
          sheet = "4_hivatkozasok",
          style = stilus_zart,
          rows = 26:max_sor,
          cols = c(1:14, 16:19),
          gridExpand = TRUE,
          stack = TRUE
        )
        
        # A kívánt oszlop visszazárása
        addStyle(
          wb = wb,
          sheet = "4_hivatkozasok",
          style = stilus_zart,
          rows = 1:21,
          cols = 1:5,
          gridExpand = TRUE,
          stack = TRUE
        )
        
        # Lapvédelem aktiválása
        protectWorksheet(wb = wb, sheet = "4_hivatkozasok", protect = TRUE)
        
        
        # --- MENTÉS ---
        log_info("[{session_id}] Excel munkafüzet mentése")
        
        # FIGYELEM: Itt a 'file' helyett a 'temp_xlsx'-be mentünk!
        saveWorkbook(wb, temp_xlsx, overwrite = TRUE)
        
        log_info("[{session_id}] SIKERES GENERÁLÁS: {temp_xlsx}")
        log_text(paste("Kész! Szerző:", name))
        
        # Ha minden sikeres, értesítjük a UI-t, hogy kész a fájl
        ready_file_path(temp_xlsx)
        
      }, error = function(e) {
        # (Ide jön az eredeti error handling részlet változatlanul)
        tb <- rlang::trace_back()
        log_error("[{session_id}] HIBA TÖRTÉNT!")
        log_error("[{session_id}] Üzenet: {conditionMessage(e)}")
        walk(format(tb), ~log_error(skip_formatter(paste("  ", .x))))
        log_text(paste("Hiba:", conditionMessage(e)))
        showModal(modalDialog(title = "Hiba történt", p(conditionMessage(e))))
      })
    })
  })
  
  # 3. A letöltés gomb csak akkor jelenik meg, ha a ready_file_path() már nem NULL
  output$downloadBtnUI <- renderUI({
    req(ready_file_path())
    downloadButton("downloadData", "Kész Excel letöltése", class = "btn-success btn-lg w-100")
  })
  
  # 4. Maga a letöltés. Ez már villámgyors lesz, mert nem itt számolunk, csak átadjuk a fájlt!
  output$downloadData <- downloadHandler(
    filename = function() {
      paste0("mtmt_output_", input$mtmtid, "_", substr(Sys.time(),1,10), ".xlsx")
    },
    content = function(file) {
      # A generált temp fájlt átmásoljuk a böngésző által várt outputba
      file.copy(ready_file_path(), file)
    }
  )
}

shinyApp(ui, server)
