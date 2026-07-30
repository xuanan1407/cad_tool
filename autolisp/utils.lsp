;; utils.lsp
;; Utility functions for CAD Column Inspector Pro
;; Copyright (c) 2026 Tran Xuan An

;; ========================================
;; STRING UTILITIES
;; ========================================

(defun str-split (str delimiter / pos result)
  "Split string by delimiter"
  (setq result '())
  (while (setq pos (vl-string-search delimiter str))
    (setq result (append result (list (substr str 1 pos))))
    (setq str (substr str (+ pos (strlen delimiter) 1)))
  )
  (append result (list str))
)

(defun str-clean (str / clean)
  "Clean string - remove spaces and special characters"
  (if str
    (progn
      (setq clean (vl-string-translate " " "_" str))
      (setq clean (vl-string-translate "-" "_" clean))
      clean
    )
    "UNKNOWN"
  )
)

(defun str-pad-zeros (num width / str)
  "Pad number with leading zeros"
  (setq str (itoa num))
  (while (< (strlen str) width)
    (setq str (strcat "0" str))
  )
  str
)

;; ========================================
;; FILE UTILITIES
;; ========================================

(defun get-result-file (/ dwg_path)
  "Get full path to result JSON file"
  (setq dwg_path (getvar "DWGPREFIX"))
  (strcat dwg_path *COL_DATA_FILENAME*)
)

(defun file-exists-p (filepath)
  "Check if file exists"
  (if (findfile filepath) T nil)
)

(defun delete-entities (ent_list / ent)
  "Delete list of entities safely"
  (foreach ent ent_list
    (if (and ent (entget ent))
      (entdel ent)
    )
  )
)

;; ========================================
;; SHAPE ID UTILITIES
;; ========================================

(defun generate-shape-id (name count / clean_name count_str)
  "Generate unique shape ID: NAME_001, NAME_002, etc."
  (setq clean_name (str-clean name))
  (setq count_str (str-pad-zeros count 3))
  (strcat clean_name "_" count_str)
)

;; ========================================
;; PYTHON UTILITIES
;; ========================================

(defun find-python-executable (/ test_path)
  "Find Python executable from common locations"
  (setq test_path nil)
  
  (foreach path *COL_PYTHON_PATHS*
    (if (and (null test_path) (findfile path))
      (setq test_path path)
    )
  )
  
  test_path
)

(defun call-python-processor (json_file / python_path python_exe)
  "Call Python processor with JSON file"
  (princ "\nCalling Python processor...")
  (setq python_path (findfile *COL_PYTHON_SCRIPT*))
  
  (if python_path
    (progn
      (setq python_exe (find-python-executable))
      
      (if python_exe
        (progn
          (startapp python_exe (strcat "\"" python_path "\" \"" json_file "\""))
          (princ "\n✓ Python processor started successfully!")
          (princ "\n✓ Check the output folder for Excel file.")
          T
        )
        (progn
          (princ "\n⚠ Warning: Python not found. Please process manually.")
          nil
        )
      )
    )
    (progn
      (princ (strcat "\n⚠ Warning: " *COL_PYTHON_SCRIPT* " not found."))
      (princ (strcat "\nData saved to: " json_file))
      (princ "\nYou can process this file manually.")
      nil
    )
  )
)

;; ========================================
;; TEXT SEARCH UTILITIES
;; ========================================

(defun find-nearest-text (center_pt / ss i ent ent_type text_string text_pt distance min_distance nearest_text)
  "Find nearest TEXT or MTEXT to a point"
  (setq min_distance 1e10)
  (setq nearest_text nil)
  
  ;; Get all text entities in drawing
  (setq ss (ssget "_X" '((0 . "TEXT,MTEXT"))))
  
  (if ss
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq ent_type (cdr (assoc 0 (entget ent))))
        
        ;; Get text content
        (setq text_string (cdr (assoc 1 (entget ent))))
        
        ;; Get insertion point
        (setq text_pt (cdr (assoc 10 (entget ent))))
        
        ;; Calculate 2D distance
        (setq distance (sqrt (+ 
          (* (- (car text_pt) (car center_pt)) (- (car text_pt) (car center_pt)))
          (* (- (cadr text_pt) (cadr center_pt)) (- (cadr text_pt) (cadr center_pt)))
        )))
        
        ;; Update if closer
        (if (< distance min_distance)
          (progn
            (setq min_distance distance)
            (setq nearest_text text_string)
          )
        )
        
        (setq i (1+ i))
      )
    )
  )
  
  nearest_text
)

;; ========================================
;; UI UTILITIES
;; ========================================

(defun print-header (title)
  "Print formatted header"
  (princ "\n")
  (princ (strcat "\n=== " title " ==="))
)

(defun print-success (message)
  "Print success message"
  (princ (strcat "\n✓ " message))
)

(defun print-error (message)
  "Print error message"
  (princ (strcat "\n✗ " message))
)

(defun print-warning (message)
  "Print warning message"
  (princ (strcat "\n⚠ " message))
)

;; ========================================
;; VISUAL HIGHLIGHT UTILITIES
;; ========================================

(defun pause-ms (milliseconds / start_time)
  "Pause for specified milliseconds"
  (setq start_time (getvar "MILLISECS"))
  (while (< (- (getvar "MILLISECS") start_time) milliseconds)
    (grread t)  ; Non-blocking read
  )
)

(defun highlight-shape-permanent (shape / shape_type points centroid radius pt1 pt2 i offset_points highlight_ent)
  "Draw permanent highlight around shape (does not auto-delete)"
  (setq shape_type (cdr (assoc "type" shape)))
  (setq points (cdr (assoc "points" shape)))
  (setq centroid (cdr (assoc "centroid" shape)))
  
  (cond
    ;; Circle - draw circle with radius using entmake
    ((= shape_type "circle")
     (setq radius (cdr (assoc "radius" shape)))
     (if (and radius centroid)
       (progn
         ;; Ensure lineweight display is on
         (setvar "LWDISPLAY" 1)
         
         ;; Create yellow circle entity directly
         (entmake 
           (list
             (cons 0 "CIRCLE")
             (cons 10 (list (car centroid) (cadr centroid) 0.0))
             (cons 40 (* radius 1.05))  ; radius (5% larger)
             (cons 62 2)  ; color = yellow
             (cons 370 100)  ; lineweight = 1.00mm
           )
         )
         (setq highlight_ent (entlast))
       )
     )
    )
    
    ;; Polygon/Polyline - draw offset polyline
    ((or (= shape_type "polygon") (= shape_type "lwpolyline") (= shape_type "polyline"))
     (if points
       (progn
         ;; Ensure lineweight display is on
         (setvar "LWDISPLAY" 1)
         
         ;; Calculate offset points (slightly outward from centroid)
         (setq offset_points '())
         (foreach pt points
           (setq pt1 (list (car pt) (cadr pt) (if (caddr pt) (caddr pt) 0.0)))
           ;; Offset 5% outward from centroid
           (setq pt2 (list
             (+ (car pt1) (* (- (car pt1) (car centroid)) 0.05))
             (+ (cadr pt1) (* (- (cadr pt1) (cadr centroid)) 0.05))
             (caddr pt1)
           ))
           (setq offset_points (append offset_points (list pt2)))
         )
         
         ;; Draw yellow polyline
         (command "_.PLINE")
         (foreach pt offset_points
           (command pt)
         )
         (command "_C")  ;; Close
         
         ;; Set to yellow and thicker lineweight
         (setq highlight_ent (entlast))
         (if highlight_ent
           (command "_.CHPROP" highlight_ent "" "_C" "2" "_LW" "100" "")
         )
       )
     )
    )
  )
  
  highlight_ent  ; Return entity handle
)

(defun highlight-shape (shape / shape_type points centroid radius pt1 pt2 i offset_points highlight_ent)
  "Draw highlight around shape and flash for 1 second"
  (setq shape_type (cdr (assoc "type" shape)))
  (setq points (cdr (assoc "points" shape)))
  (setq centroid (cdr (assoc "centroid" shape)))
  
  (princ (strcat "\n✓ Shape found: " (cdr (assoc "name" shape))))
  (princ (strcat "\n  Type: " shape_type))
  (princ (strcat "\n  Area: " (rtos (cdr (assoc "area" shape)) 2 2) " sq units"))
  
  (cond
    ;; Circle - draw circle with radius using entmake
    ((= shape_type "circle")
     (setq radius (cdr (assoc "radius" shape)))
     (if (and radius centroid)
       (progn
         ;; Ensure lineweight display is on
         (setvar "LWDISPLAY" 1)
         
         ;; Create yellow circle entity directly (no interactive command)
         (entmake 
           (list
             (cons 0 "CIRCLE")
             (cons 10 (list (car centroid) (cadr centroid) 0.0))  ; center point
             (cons 40 (* radius 1.05))  ; radius (5% larger)
             (cons 62 2)  ; color = yellow
             (cons 370 100)  ; lineweight = 1.00mm
           )
         )
         
         ;; Get the created entity
         (setq highlight_ent (entlast))
         (if highlight_ent
           (progn
             (princ "\n✓ Yellow highlight drawn around circle")
             
             ;; Zoom to shape
             (command "_.ZOOM" "_C" (list (car centroid) (cadr centroid) 0.0) (* radius 3))
             (command "_.REDRAW")
             
             ;; Flash for 1 second then delete
             (princ "\n  Flashing for 1 second...")
             (pause-ms 1000)
             (entdel highlight_ent)
             (command "_.REDRAW")
             (princ "\n✓ Highlight removed")
           )
         )
       )
     )
    )
    
    ;; Polygon/Polyline - draw offset polyline
    ((or (= shape_type "polygon") (= shape_type "lwpolyline") (= shape_type "polyline"))
     (if points
       (progn
         ;; Ensure lineweight display is on
         (setvar "LWDISPLAY" 1)
         
         ;; Calculate offset points (slightly outward from centroid)
         (setq offset_points '())
         (foreach pt points
           (setq pt1 (list (car pt) (cadr pt) (if (caddr pt) (caddr pt) 0.0)))
           ;; Offset 5% outward from centroid
           (setq pt2 (list
             (+ (car pt1) (* (- (car pt1) (car centroid)) 0.05))
             (+ (cadr pt1) (* (- (cadr pt1) (cadr centroid)) 0.05))
             (caddr pt1)
           ))
           (setq offset_points (append offset_points (list pt2)))
         )
         
         ;; Draw yellow polyline
         (command "_.PLINE")
         (foreach pt offset_points
           (command pt)
         )
         (command "_C")  ;; Close
         
         ;; Set to yellow and thicker lineweight
         (setq highlight_ent (entlast))
         (if highlight_ent
           (progn
             (command "_.CHPROP" highlight_ent "" "_C" "2" "_LW" "100" "")
             (princ "\n✓ Yellow highlight drawn around polygon")
             
             ;; Zoom to shape
             (if centroid
               (command "_.ZOOM" "_C" (list (car centroid) (cadr centroid) 0.0) 500)
             )
             (command "_.REDRAW")
             
             ;; Flash for 1 second then delete
             (princ "\n  Flashing for 1 second...")
             (pause-ms 1000)
             (entdel highlight_ent)
             (command "_.REDRAW")
             (princ "\n✓ Highlight removed")
           )
         )
       )
     )
    )
  )
)

(princ "\n✓ Utilities loaded")
(princ)
