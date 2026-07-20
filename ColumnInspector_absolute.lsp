;; ColumnInspector_absolute.lsp
;; Version with absolute path
;; EDIT the path below to match your system!

(princ "\n=== CAD Column Inspector - Absolute Path Loader ===\n")

;; ========================================
;; CONFIGURE THIS PATH!
;; ========================================
;; Set to your project folder
(setq *COL_PROJECT_PATH* "C:/Users/Admin/Desktop/Work/cad_tool/")

;; ========================================
;; LOAD MODULES
;; ========================================

(princ "\nLoading config...")
(load (strcat *COL_PROJECT_PATH* "autolisp/config.lsp"))

(princ "\nLoading utils...")
(load (strcat *COL_PROJECT_PATH* "autolisp/utils.lsp"))

(princ "\nLoading geometry...")
(load (strcat *COL_PROJECT_PATH* "autolisp/geometry.lsp"))

(princ "\nLoading detection...")
(load (strcat *COL_PROJECT_PATH* "autolisp/detection.lsp"))

(princ "\nLoading io...")
(load (strcat *COL_PROJECT_PATH* "autolisp/io.lsp"))

(princ "\nLoading commands...")
(load (strcat *COL_PROJECT_PATH* "autolisp/commands.lsp"))

(princ "\n\n=== ALL MODULES LOADED ===")
(princ "\nCommands: COLINSPECT, COLSCAN, COLLIST\n")
(princ)
