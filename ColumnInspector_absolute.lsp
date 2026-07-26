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

(princ "\nLoading ui...")
(load (strcat *COL_PROJECT_PATH* "autolisp/ui.lsp"))

(princ "\n\n=== ALL MODULES LOADED ===")
(princ "\n╔═══════════════════════════════════════════╗")
(princ "\n║  CAD Column Inspector Pro - Ready!       ║")
(princ "\n╚═══════════════════════════════════════════╝")
(princ "\n")
(princ "\n🎯 MAIN COMMAND:")
(princ "\n   COLUI - Open UI Dialog (Recommended!)")
(princ "\n")
(princ "\n📝 COMMAND LINE ALTERNATIVES:")
(princ "\n   COLINSPECT  - Manual polygon selection")
(princ "\n   COLSCAN     - Auto-scan area for shapes")
(princ "\n   COLFINDPOLY - Find shape by ID")
(princ "\n   COLLIST     - List all shapes")
(princ "\n   COLCLEAR    - Clear all data")
(princ "\n")
(princ)
