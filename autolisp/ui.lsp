;; ui.lsp
;; UI Dialog handling for CAD Column Inspector Pro
;; Copyright (c) 2026 Tran Xuan An

;; ========================================
;; UI DIALOG HANDLER
;; ========================================

(defun show-col-inspector-ui (/ dcl_id status dcl_file)
  ;; Use absolute path for DCL file
  (setq dcl_file "C:/Users/Admin/Desktop/Work/cad_tool/autolisp/col_inspector_ui.dcl")
  
  ;; Check if file exists
  (if (not (findfile dcl_file))
    (progn
      (print-error "Cannot find col_inspector_ui.dcl file!")
      (princ (strcat "\nLooking for: " dcl_file))
      (setq dcl_id nil)
    )
    (progn
      (princ (strcat "\nLoading DCL from: " dcl_file))
      (setq dcl_id (load_dialog dcl_file))
    )
  )
  
  (if (or (not dcl_id) (< dcl_id 0))
    (progn
      (print-error "Cannot load dialog file!")
      (princ (strcat "\nTried path: " (if dcl_file dcl_file "nil")))
      nil
    )
    (progn
      (if (not (new_dialog "col_inspector" dcl_id))
        (progn
          (print-error "Cannot load dialog!")
          (unload_dialog dcl_id)
          nil
        )
        (progn
          ;; Set up button actions
          (action_tile "btn_inspect" "(done_dialog 1)")
          (action_tile "btn_scan" "(done_dialog 2)")
          (action_tile "btn_clear" "(done_dialog 3)")
          (action_tile "btn_export" "(done_dialog 4)")
          (action_tile "btn_help" "(done_dialog 5)")
          (action_tile "accept" "(done_dialog 0)")
          
          ;; Show dialog and get result
          (setq status (start_dialog))
          (unload_dialog dcl_id)
          
          ;; Process result
          (cond
            ((= status 1)  ; Manual Inspect
             (princ "\n")
             (C:COLINSPECT)
             (show-col-inspector-ui))  ; Reopen UI after action
            
            ((= status 2)  ; Auto Scan
             (princ "\n")
             (C:COLSCAN)
             (show-col-inspector-ui))
            
            ((= status 3)  ; Clear Data
             (princ "\n")
             (C:COLCLEAR)
             (show-col-inspector-ui))
            
            ((= status 4)  ; Export to Excel
             (princ "\n")
             (export-to-excel))
            
            ((= status 5)  ; Help
             (show-help-dialog)
             (show-col-inspector-ui))
            
            ((= status 0)  ; Close
             (princ "\nDialog closed."))
          )
        )
      )
    )
  )
  (princ)
)

;; ========================================
;; EXPORT TO EXCEL FUNCTION
;; ========================================

(defun export-to-excel (/ result_file)
  (setq result_file (get-result-file))
  
  (if (file-exists-p result_file)
    (progn
      (print-header "Export to Excel")
      (call-python-processor result_file)
      (print-success "Export completed! Check the Excel file.")
    )
    (print-warning "No data found. Run COLINSPECT or COLSCAN first.")
  )
  (princ)
)

;; ========================================
;; HELP DIALOG
;; ========================================

(defun show-help-dialog ()
  (print-header "CAD Column Inspector Pro - Help")
  (princ "\n")
  (princ "\nMAIN OPERATIONS:")
  (princ "\n  • Manual Inspect: Select points to create polygon manually")
  (princ "\n  • Auto Scan: Automatically scan area for shapes")
  (princ "\n")
  (princ "\nDATA MANAGEMENT:")
  (princ "\n  • Clear Data: Delete all saved polygon data")
  (princ "\n  • Export Excel: Generate Excel report from data")
  (princ "\n")
  (princ "\nCOMMAND LINE:")
  (princ "\n  • COLUI - Open this UI dialog")
  (princ "\n  • COLINSPECT - Manual polygon selection")
  (princ "\n  • COLSCAN - Auto-scan area")
  (princ "\n  • COLCLEAR - Clear all data")
  (princ "\n  • COLFINDPOLY - Find single shape by exact ID")
  (princ "\n  • COLFINDGROUP - Find multiple shapes by prefix (NEW!)")
  (princ "\n  • COLLIST - List all shapes")
  (princ "\n")
  (princ "\nPress Enter to continue...")
  (getstring)
  (princ)
)

;; ========================================
;; COMMAND: COLUI
;; Open UI Dialog
;; ========================================

(defun C:COLUI ()
  (show-col-inspector-ui)
  (princ)
)

(princ "\n✓ UI module loaded")
(princ)
