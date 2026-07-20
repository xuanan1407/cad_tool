;; config.lsp
;; Configuration constants for CAD Column Inspector Pro
;; Copyright (c) 2026 Tran Xuan An

;; ========================================
;; FILE PATHS
;; ========================================
(setq *COL_DATA_FILENAME* "column_data.json")
(setq *COL_PYTHON_SCRIPT* "main.py")

;; ========================================
;; PROCESSING SETTINGS
;; ========================================
(setq *COL_DEFAULT_TOLERANCE* 20.0)  ; Default minimum distance between shapes (mm)
(setq *COL_AREA_TOLERANCE* 0.1)      ; Area comparison tolerance for consistency checks

;; ========================================
;; VISUAL SETTINGS
;; ========================================
(setq *COL_PREVIEW_COLOR* 1)         ; Red color for preview lines
(setq *COL_PREVIEW_WIDTH* 2.0)       ; Line width for preview
(setq *COL_HIGHLIGHT_COLOR* 2)       ; Yellow color for shape highlighting

;; ========================================
;; PYTHON EXECUTABLE SEARCH PATHS
;; ========================================
(setq *COL_PYTHON_PATHS* '(
  "C:\\Python312\\python.exe"
  "C:\\Python311\\python.exe"
  "C:\\Python310\\python.exe"
  "C:\\Python39\\python.exe"
  "C:\\Program Files\\Python312\\python.exe"
  "C:\\Program Files\\Python311\\python.exe"
  "C:\\Program Files\\Python310\\python.exe"
  "C:\\Program Files\\Python39\\python.exe"
  "python.exe"  ; Try PATH
))

;; ========================================
;; ENTITY FILTERS FOR AUTO-SCAN
;; ========================================
(setq *COL_SCAN_FILTER* '(
  (-4 . "<OR")
    (0 . "LWPOLYLINE")
    (0 . "POLYLINE")
    (0 . "CIRCLE")
  (-4 . "OR>")
))

;; ========================================
;; GLOBAL VARIABLES
;; ========================================
(setq *COLFIND_SHAPEID* nil)  ; Used by reverse lookup feature

(princ "\n✓ Configuration loaded")
(princ)
