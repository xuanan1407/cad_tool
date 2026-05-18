;; ColumnInspector.lsp
;; CAD Column Inspector Pro - AutoLISP Plugin
;; Copyright (c) 2026 Tran Xuan An

;; Load this file in AutoCAD: (load "ColumnInspector.lsp")
;; Run command: COLINSPECT

(defun C:COLINSPECT (/ pt_list pt count area centroid result_file python_path python_exe confirm)
  (princ "\n=== CAD Column Inspector Pro ===")
  (princ "\nSelect points to create polygon (Press Enter to finish)")
  
  ;; Initialize variables
  (setq pt_list '())
  (setq count 0)
  
  ;; Get points from user
  (while (setq pt (getpoint 
                    (if (null pt_list)
                      "\nSelect first point: "
                      "\nSelect next point (or press Enter to finish): "
                    )))
    (setq pt_list (append pt_list (list pt)))
    (setq count (1+ count))
    (princ (strcat "\nPoint " (itoa count) " selected: " (vl-princ-to-string pt)))
  )
  
  ;; Check if we have at least 3 points
  (if (< (length pt_list) 3)
    (progn
      (princ "\nError: Need at least 3 points to create a polygon!")
      (princ)
    )
    (progn
      ;; Draw temporary polygon for preview
      (command "_.PLINE")
      (foreach pt pt_list
        (command pt)
      )
      (command "_C") ;; Close polyline
      
      ;; Calculate polygon area and centroid
      (setq area (calculate-polygon-area pt_list))
      (setq centroid (calculate-centroid pt_list))
      
      ;; Find nearest text to centroid
      (setq nearest_text (find-nearest-text centroid))
      
      ;; Display info
      (princ (strcat "\n\nPolygon created with " (itoa count) " points"))
      (if nearest_text
        (princ (strcat "\nName: " nearest_text))
        (princ "\nName: [No text found]")
      )
      (princ (strcat "\nArea: " (rtos area 2 2) " sq units"))
      (princ (strcat "\nCentroid: " (vl-princ-to-string centroid)))
      
      ;; Ask user to confirm
      (initget "Yes No")
      (setq confirm (getkword "\nExport this polygon data? [Yes/No] <Yes>: "))
      
      (if (or (null confirm) (= confirm "Yes"))
        (progn
          ;; Save data to JSON file
          (setq result_file (strcat (getvar "DWGPREFIX") "column_data.json"))
          (append-polygon-data pt_list area centroid nearest_text result_file)
          
          ;; Delete the temporary polyline after saving
          (command "_.ERASE" "L" "")
          
          ;; Call Python script to process
          (princ "\nCalling Python processor...")
          (setq python_path (findfile "python_processor.py"))
          
          (if python_path
            (progn
              ;; Find Python executable
              (setq python_exe (find-python-executable))
              
              (if python_exe
                (progn
                  (startapp python_exe (strcat "\"" python_path "\" \"" result_file "\""))
                  (princ "\nPython processor started successfully!")
                  (princ "\nCheck the output folder for Excel file.")
                )
                (princ "\nWarning: Python not found. Please process manually.")
              )
            )
            (progn
              (princ "\nWarning: python_processor.py not found.")
              (princ (strcat "\nData saved to: " result_file))
              (princ "\nYou can process this file manually.")
            )
          )
        )
        (progn
          ;; User chose No, delete the temporary polyline
          (command "_.ERASE" "L" "")
          (princ "\nOperation cancelled.")
        )
      )
      (princ)
    )
  )
  (princ)
)

;; Helper function: Calculate polygon area using shoelace formula
(defun calculate-polygon-area (pt_list / area i n x1 y1 x2 y2)
  (setq area 0.0)
  (setq n (length pt_list))
  (setq i 0)
  
  (repeat n
    (setq x1 (car (nth i pt_list)))
    (setq y1 (cadr (nth i pt_list)))
    (setq x2 (car (nth (if (= i (- n 1)) 0 (+ i 1)) pt_list)))
    (setq y2 (cadr (nth (if (= i (- n 1)) 0 (+ i 1)) pt_list)))
    
    (setq area (+ area (* x1 y2)))
    (setq area (- area (* x2 y1)))
    
    (setq i (1+ i))
  )
  
  (abs (/ area 2.0))
)

;; Helper function: Calculate centroid
(defun calculate-centroid (pt_list / sum_x sum_y count)
  (setq sum_x 0.0)
  (setq sum_y 0.0)
  (setq count (length pt_list))
  
  (foreach pt pt_list
    (setq sum_x (+ sum_x (car pt)))
    (setq sum_y (+ sum_y (cadr pt)))
  )
  
  (list (/ sum_x count) (/ sum_y count))
)

;; Helper function: Append polygon data to JSON array
(defun append-polygon-data (pt_list area centroid name filename / old_content lines new_content polygon_count file i j pt)
  ;; Check if file exists and read old content
  (setq old_content '())
  (setq polygon_count 0)
  
  (if (findfile filename)
    (progn
      ;; Read existing file content line by line
      (setq file (open filename "r"))
      (if file
        (progn
          (while (setq line (read-line file))
            (setq old_content (append old_content (list line)))
          )
          (close file)
          
          ;; Count existing polygons by counting lines with "  {"
          (foreach line old_content
            (if (= line "  {")
              (setq polygon_count (1+ polygon_count))
            )
          )
        )
      )
    )
  )
  
  ;; Now write the new file
  (setq file (open filename "w"))
  
  (if file
    (progn
      (if (= polygon_count 0)
        ;; First polygon - create new array
        (progn
          (write-line "[" file)
        )
        ;; Existing polygons - copy old content without closing ]
        (progn
          (setq i 0)
          (foreach line old_content
            (if (< i (- (length old_content) 1))  ;; Skip last line (the ])
              (progn
                ;; If this is the last "  }" before ], add comma
                (if (and (= line "  }") (= i (- (length old_content) 2)))
                  (write-line "  }," file)
                  (write-line line file)
                )
              )
            )
            (setq i (1+ i))
          )
        )
      )
      
      ;; Write the new polygon
      (write-line "  {" file)
      (write-line "    \"type\": \"polygon\"," file)
      (write-line (strcat "    \"name\": \"" (if name name "Unknown") "\",") file)
      (write-line (strcat "    \"point_count\": " (itoa (length pt_list)) ",") file)
      (write-line (strcat "    \"area\": " (rtos area 2 6) ",") file)
      
      ;; Centroid
      (write-line (strcat "    \"centroid\": [" 
                         (rtos (car centroid) 2 6) ", " 
                         (rtos (cadr centroid) 2 6) 
                         "],") file)
      
      ;; Points
      (write-line "    \"points\": [" file)
      (setq j 0)
      (foreach pt pt_list
        (write-line (strcat "      [" 
                           (rtos (car pt) 2 6) ", " 
                           (rtos (cadr pt) 2 6) ", "
                           (rtos (caddr pt) 2 6) "]"
                           (if (< j (- (length pt_list) 1)) "," "")) file)
        (setq j (1+ j))
      )
      (write-line "    ]," file)
      
      ;; Timestamp and drawing
      (write-line (strcat "    \"timestamp\": \"" (rtos (getvar "CDATE") 2 0) "\",") file)
      (write-line (strcat "    \"drawing\": \"" (getvar "DWGNAME") "\"") file)
      
      ;; Close polygon object (no comma on last one)
      (write-line "  }" file)
      
      ;; Close array
      (write-line "]" file)
      
      (close file)
      
      ;; Update count
      (setq polygon_count (1+ polygon_count))
      (princ (strcat "\n✓ Polygon #" (itoa polygon_count) " saved to: " filename))
      (princ (strcat "\nTotal polygons: " (itoa polygon_count)))
    )
    (princ "\n✗ Error: Cannot create output file!")
  )
)

;; Helper function: Find Python executable
(defun find-python-executable (/ python_paths test_path)
  (setq python_paths '(
    "C:\\Python312\\python.exe"
    "C:\\Python311\\python.exe"
    "C:\\Python310\\python.exe"
    "C:\\Program Files\\Python312\\python.exe"
    "C:\\Program Files\\Python311\\python.exe"
    "python.exe"  ;; Try PATH
  ))
  
  ;; Try to find python in common locations
  (foreach path python_paths
    (if (and (null test_path) (findfile path))
      (setq test_path path)
    )
  )
  
  test_path
)

;; Helper function: Find nearest text to a point
(defun find-nearest-text (center_pt / ss i ent ent_type text_string text_pt distance min_distance nearest_text)
  (setq min_distance 1e10)  ;; Very large number
  (setq nearest_text nil)
  
  ;; Get all entities in the drawing
  (setq ss (ssget "_X" '((0 . "TEXT,MTEXT"))))
  
  (if ss
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq ent_type (cdr (assoc 0 (entget ent))))
        
        ;; Get text string
        (setq text_string (cdr (assoc 1 (entget ent))))
        
        ;; Get text insertion point
        (if (= ent_type "TEXT")
          (setq text_pt (cdr (assoc 10 (entget ent))))
          ;; For MTEXT, use insertion point
          (setq text_pt (cdr (assoc 10 (entget ent))))
        )
        
        ;; Calculate distance from text to centroid (2D distance)
        (setq distance (sqrt (+ 
          (* (- (car text_pt) (car center_pt)) (- (car text_pt) (car center_pt)))
          (* (- (cadr text_pt) (cadr center_pt)) (- (cadr text_pt) (cadr center_pt)))
        )))
        
        ;; Update if this is closer
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

;; Auto-load message
(princ "\n=== CAD Column Inspector Pro Loaded ===")
(princ "\nType COLINSPECT to start selecting polygon points")
(princ "\nType COLCLEAR to clear all saved polygons and start fresh")
(princ "\n")
(princ)

;; Command: Clear all polygon data
(defun C:COLCLEAR (/ result_file confirm)
  (princ "\n=== Clear Polygon Data ===")
  
  ;; Ask user to confirm
  (initget "Yes No")
  (setq confirm (getkword "\nDelete all saved polygons and start fresh? [Yes/No] <No>: "))
  
  (if (= confirm "Yes")
    (progn
      (setq result_file (strcat (getvar "DWGPREFIX") "column_data.json"))
      
      ;; Delete the file if it exists
      (if (findfile result_file)
        (progn
          (vl-file-delete result_file)
          (princ "\n✓ All polygon data cleared!")
          (princ "\nYou can now start selecting new polygons with COLINSPECT")
        )
        (princ "\nNo polygon data file found.")
      )
    )
    (princ "\nOperation cancelled.")
  )
  (princ)
)
