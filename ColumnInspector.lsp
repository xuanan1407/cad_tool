;; ColumnInspector.lsp
;; CAD Column Inspector Pro - AutoLISP Plugin
;; Copyright (c) 2026 Tran Xuan An

;; Load this file in AutoCAD: (load "ColumnInspector.lsp")
;; Run command: COLINSPECT

(defun C:COLINSPECT (/ pt_list pt count area centroid result_file python_path python_exe confirm redline_list last_ent ent_data prev_pt nearest_text current_shape history_shapes global_shapes error_count)
  (princ "\n=== CAD Column Inspector Pro ===")
  (princ "\nSelect points to create polygon (Press Enter to finish)")
  
  ;; Initialize variables
  (setq pt_list '())
  (setq count 0)
  (setq redline_list '())  ;; List to store red line entities
  
  ;; Get points from user
  (while (setq pt (if (null pt_list)
                    ;; First point - no base point
                    (getpoint "\nSelect first point: ")
                    ;; Subsequent points - use last point as base
                    (getpoint (car (reverse pt_list)) "\nSelect next point (or press Enter to finish): ")
                  ))
    (setq pt_list (append pt_list (list pt)))
    (setq count (1+ count))
    (princ (strcat "\nPoint " (itoa count) " selected: " (vl-princ-to-string pt)))
    
    ;; Draw red line from previous point to current point
    (if (> count 1)
      (progn
        (setq prev_pt (nth (- count 2) pt_list))
        
        ;; Draw THICK red polyline
        (command "_.PLINE" prev_pt "W" "2" "2" pt "")
        
        ;; Get the line entity and change to RED
        (setq last_ent (entlast))
        (if last_ent
          (progn
            (setq ent_data (entget last_ent))
            
            ;; Add or change color to RED (1)
            (if (assoc 62 ent_data)
              (setq ent_data (subst (cons 62 1) (assoc 62 ent_data) ent_data))
              (setq ent_data (append ent_data (list (cons 62 1))))
            )
            
            (entmod ent_data)
            (entupd last_ent)
            
            ;; Store entity name for later cleanup
            (setq redline_list (append redline_list (list last_ent)))
          )
        )
      )
    )
  )
  
  ;; Check if we have at least 3 points
  (if (< (length pt_list) 3)
    (progn
      (princ "\nError: Need at least 3 points to create a polygon!")
      
      ;; Delete red lines
      (foreach ent redline_list
        (if (and ent (entget ent))
          (entdel ent)
        )
      )
      (princ)
    )
    (progn
      ;; Delete red lines before drawing final polygon
      (foreach ent redline_list
        (if (and ent (entget ent))
          (entdel ent)
        )
      )
      
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
      (if (null nearest_text) (setq nearest_text "[No text found]"))
      
      ;; ==========================================================
      ;; NEW LOGIC: LOAD HISTORY & CROSS-CHECK AREA CONSISTENCY
      ;; ==========================================================
      (setq result_file (strcat (getvar "DWGPREFIX") "column_data.json"))
      (setq history_shapes '())
      
      ;; 1. Load historical data if file exists
      (if (findfile result_file)
        (setq history_shapes (load-shapes-from-json result_file))
      )
      
      ;; 2. Format the current inspected shape into standard list structure
      (setq current_shape (list (cons "name" nearest_text) (cons "area" area)))
      
      ;; 3. Combine history data with the current inspected shape
      (setq global_shapes (append history_shapes (list current_shape)))
      
      ;; 4. Run consistency check on the combined global data
      (setq error_count (check-area-consistency global_shapes))
      ;; ==========================================================
      
      ;; Display info
      (princ (strcat "\n\nPolygon created with " (itoa count) " points"))
      (princ (strcat "\nName: " nearest_text))
      (princ (strcat "\nArea: " (rtos area 2 2) " sq units"))
      (princ (strcat "\nCentroid: " (vl-princ-to-string centroid)))
      
      ;; Ask user to confirm
      (initget "Yes No")
      ;; Smart prompt based on error status
      (if (= error_count 0)
        (setq confirm (getkword "\nExport this polygon data? [Yes/No] <Yes>: "))
        (setq confirm (getkword "\n⚠ WARNING: Data contains geometric inconsistencies! Force export and append to JSON? [Yes/No] <No>: "))
      )
      
      (if (or (and (null confirm) (= error_count 0)) (= confirm "Yes"))
        (progn
          ;; Save data to JSON file (Passing only the current newly picked data)
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

;; Helper function: Generate unique shape ID
(defun generate-shape-id (name count / clean_name count_str)
  ;; Clean name: remove spaces and special chars
  (setq clean_name (vl-string-translate " " "_" (if name name "UNKNOWN")))
  (setq clean_name (vl-string-translate "-" "_" clean_name))
  
  ;; Format count with leading zeros (3 digits: 001, 002, etc.)
  (setq count_str (itoa count))
  (while (< (strlen count_str) 3)
    (setq count_str (strcat "0" count_str))
  )
  
  (strcat clean_name "_" count_str)
)

;; Helper function: Append polygon data to JSON array
(defun append-polygon-data (pt_list area centroid name filename / old_content lines new_content polygon_count file i j pt shape_id)
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
      
      ;; Generate unique ID
      (setq shape_id (generate-shape-id name (+ polygon_count 1)))
      
      ;; Write the new polygon
      (write-line "  {" file)
      (write-line (strcat "    \"id\": \"" shape_id "\",") file)
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

;; Helper function: Check if two shapes are too close
(defun is-duplicate-shape (shape1 shape2 tolerance / centroid1 centroid2 dist)
  (setq centroid1 (cdr (assoc "centroid" shape1)))
  (setq centroid2 (cdr (assoc "centroid" shape2)))
  
  ;; Check if centroids exist
  (if (and centroid1 centroid2)
    (progn
      ;; Calculate 2D distance between centroids
      (setq dist (sqrt (+ 
        (* (- (car centroid1) (car centroid2)) (- (car centroid1) (car centroid2)))
        (* (- (cadr centroid1) (cadr centroid2)) (- (cadr centroid1) (cadr centroid2)))
      )))
      
      ;; Return true if distance is less than tolerance
      (< dist tolerance)
    )
    nil
  )
)

;; Helper function: Remove shapes that are too close
(defun remove-duplicate-shapes (shapes tolerance / unique_shapes shape is_dup unique_shape)
  (setq unique_shapes '())
  
  (foreach shape shapes
    (setq is_dup nil)
    
    ;; Check if this shape is too close to any shape in unique_shapes
    (foreach unique_shape unique_shapes
      (if (is-duplicate-shape shape unique_shape tolerance)
        (setq is_dup T)
      )
    )
    
    ;; If not too close, add to unique list
    (if (not is_dup)
      (setq unique_shapes (append unique_shapes (list shape)))
    )
  )
  
  unique_shapes
)

;; Function: Check for area inconsistencies within the same shape name
(defun check-area-consistency (shapes / checklist errors item name area existing_area)
  (setq checklist '())
  (setq errors 0)
  (princ "\n--- CHECKING SHAPE AREA CONSISTENCY ---")
  
  (foreach item shapes
    (setq name (cdr (assoc "name" item)))
    (setq area (cdr (assoc "area" item)))
    
    ;; Check if this shape name has been scanned before
    (setq existing_area (cdr (assoc name checklist)))
    
    (if existing_area
      ;; If name exists, check if area matches (allowing a minor tolerance of 0.1)
      (if (> (abs (- area existing_area)) 0.1) 
        (progn
          (princ (strcat "\n⚠ WARNING: Inconsistent area detected for group [" name "]! (" 
                         (rtos existing_area 2 2) " vs " (rtos area 2 2) ")"))
          (setq errors (1+ errors))
        )
      )
      ;; If it's a new name, store its area as the baseline reference
      (setq checklist (append checklist (list (cons name area))))
    )
  )
  
  (if (> errors 0)
    (princ (strcat "\n\n❌ ERROR: Found " (itoa errors) " geometric inconsistency/inconsistencies! Please verify the drawing.\n"))
    (princ "\n✓ Success: All shapes with the same name have consistent areas.\n")
  )
  errors ;; Returns total error count
)

;; Helper: Load shapes from your specific JSON structure back into a LISP list
(defun load-shapes-from-json (file_path / f_in line shapes current_name current_area pos_colon raw_val)
  (setq shapes '())
  (setq f_in (open file_path "r"))
  (if f_in
    (progn
      (while (setq line (read-line f_in))
        (setq line (vl-string-trim " \t," line))
        
        ;; Extract "name" value
        (if (vl-string-search "\"name\":" line)
          (progn
            (setq pos_colon (vl-string-search ":" line))
            ;; Trim brackets, quotes, and commas to get clean string name
            (setq current_name (vl-string-trim " \"\t," (substr line (+ pos_colon 2))))
          )
        )
        
        ;; Extract "area" value
        (if (vl-string-search "\"area\":" line)
          (progn
            (setq pos_colon (vl-string-search ":" line))
            (setq raw_val (vl-string-trim " \"\t," (substr line (+ pos_colon 2))))
            (setq current_area (distof raw_val))
            
            ;; Once both name and area are extracted for the object, append to list
            (if (and current_name current_area)
              (progn
                (setq shapes (append shapes (list (list (cons "name" current_name) (cons "area" current_area)))))
                ;; Reset for the next object block
                (setq current_name nil
                      current_area nil)
              )
            )
          )
        )
      )
      (close f_in)
    )
  )
  shapes ;; Returns the filtered list of history shapes
)

;; Command: Auto-scan area for shapes
(defun C:COLSCAN (/ p1 p2 ss i ent shape_data all_shapes result_file python_path python_exe original_count duplicate_count confirm tolerance_input tolerance)
  (princ "\n=== CAD Column Auto-Scanner ===")
  
  ;; Ask for tolerance distance
  (setq tolerance_input (getreal "\nEnter minimum distance between shapes (default 20.0mm): "))
  (if (null tolerance_input)
    (setq tolerance 20.0)
    (setq tolerance tolerance_input)
  )
  
  (princ (strcat "\n✓ Tolerance: " (rtos tolerance 2 2) "mm (measuring from centroid to centroid)"))
  (princ "\n✓ Any shapes closer than this will be removed (keeping the first one)")
  (princ "\nSelect area to scan for shapes...")
  
  ;; Get selection window from user
  (setq p1 (getpoint "\nSelect first corner: "))
  
  (if p1
    (progn
      (setq p2 (getcorner p1 "\nSelect opposite corner: "))
      
      (if p2
        (progn
          ;; Select all closed shapes in the window
          (setq ss (ssget "_W" p1 p2 '(
            (-4 . "<OR")
              (0 . "LWPOLYLINE")
              (0 . "POLYLINE")
              (0 . "CIRCLE")
            (-4 . "OR>")
          )))
          
          (if ss
            (progn
              (setq all_shapes '())
              (setq i 0)
              
              (princ (strcat "\n✓ Found " (itoa (sslength ss)) " shape(s) in selected area"))
              (princ "\nAnalyzing shapes...")
              
              ;; Process each entity
              (repeat (sslength ss)
                (setq ent (ssname ss i))
                (setq shape_data (analyze-entity ent))
                
                (if shape_data
                  (progn
                    (setq all_shapes (append all_shapes (list shape_data)))
                    (princ (strcat "\n  Shape #" (itoa (+ i 1)) ": " 
                                  (cdr (assoc "name" shape_data)) 
                                  " - Area: " 
                                  (rtos (cdr (assoc "area" shape_data)) 2 2)))
                  )
                )
                
                (setq i (1+ i))
              )
              
              ;; Remove shapes that are too close (This is your existing code)
              (setq original_count (length all_shapes))
              (setq all_shapes (remove-duplicate-shapes all_shapes tolerance))
              (setq duplicate_count (- original_count (length all_shapes)))
              
              ;; ==========================================================
              ;; NEW LOGIC: LOAD HISTORY & CROSS-CHECK AREA CONSISTENCY
              ;; ==========================================================
              (setq result_file (strcat (getvar "DWGPREFIX") "column_data.json"))
              (setq history_shapes '())
              
              ;; 1. Load historical data if file exists
              (if (findfile result_file)
                (setq history_shapes (load-shapes-from-json result_file))
              )
              
              ;; 2. Combine history data with newly scanned data
              (setq global_shapes (append history_shapes all_shapes))
              
              ;; 3. Run consistency check on the combined global data
              ;; (If a name has different areas between scans, it will trigger the warning)
              (setq error_count (check-area-consistency global_shapes))
              ;; ==========================================================
              
              ;; Display summary
              (princ (strcat "\n\n✓ Total shapes detected in this scan: " (itoa original_count)))
              (if (> duplicate_count 0)
                (princ (strcat "\n⚠ Removed " (itoa duplicate_count) " shape(s) within " (rtos tolerance 2 2) "mm from each other"))
              )
              (princ (strcat "\n✓ Unique shapes in this scan: " (itoa (length all_shapes))))
              
              ;; Ask to save
              (if (> (length all_shapes) 0)
                (progn
                  (initget "Yes No")
                  (setq confirm (getkword "\nExport all shapes to JSON? [Yes/No] <Yes>: "))
                  
                  (if (or (null confirm) (= confirm "Yes"))
                    (progn
                      ;; Save all shapes to JSON
                      (setq result_file (strcat (getvar "DWGPREFIX") "column_data.json"))
                      (save-all-shapes all_shapes result_file)
                      
                      ;; Call Python processor
                      (princ "\nCalling Python processor...")
                      (setq python_path (findfile "python_processor.py"))
                      
                      (if python_path
                        (progn
                          (setq python_exe (find-python-executable))
                          
                          (if python_exe
                            (progn
                              (startapp python_exe (strcat "\"" python_path "\" \"" result_file "\""))
                              (princ "\n✓ Python processor started!")
                              (princ "\n✓ Check output folder for Excel file.")
                            )
                            (princ "\nWarning: Python not found. Data saved to JSON.")
                          )
                        )
                        (princ "\nWarning: python_processor.py not found. Data saved to JSON.")
                      )
                    )
                    (princ "\nOperation cancelled.")
                  )
                )
                (princ "\nNo valid shapes found!")
              )
            )
            (princ "\nNo shapes found in selected area!")
          )
        )
        (princ "\nSelection cancelled.")
      )
    )
    (princ "\nSelection cancelled.")
  )
  (princ)
)

;; Helper function: Analyze entity and extract data
(defun analyze-entity (ent / ent_data ent_type obj area centroid pt_list name shape_type radius center)
  (setq ent_data (entget ent))
  (setq ent_type (cdr (assoc 0 ent_data)))
  (setq obj (vlax-ename->vla-object ent))
  
  (cond
    ;; Handle LWPOLYLINE
    ((= ent_type "LWPOLYLINE")
     (if (= (vlax-get-property obj 'Closed) :vlax-true)
       (progn
         (setq area (vlax-get-property obj 'Area))
         (setq pt_list (get-lwpolyline-points ent))
         (setq centroid (calculate-centroid pt_list))
         (setq name (find-nearest-text centroid))
         (setq shape_type "lwpolyline")
         
         (list
           (cons "type" shape_type)
           (cons "name" (if name name "Unknown"))
           (cons "area" area)
           (cons "centroid" centroid)
           (cons "points" pt_list)
           (cons "point_count" (length pt_list))
         )
       )
       nil  ;; Not closed
     )
    )
    
    ;; Handle POLYLINE
    ((= ent_type "POLYLINE")
     (if (= (logand (cdr (assoc 70 ent_data)) 1) 1)  ;; Check closed flag
       (progn
         (setq area (vlax-get-property obj 'Area))
         (setq pt_list (get-polyline-points ent))
         (setq centroid (calculate-centroid pt_list))
         (setq name (find-nearest-text centroid))
         (setq shape_type "polyline")
         
         (list
           (cons "type" shape_type)
           (cons "name" (if name name "Unknown"))
           (cons "area" area)
           (cons "centroid" centroid)
           (cons "points" pt_list)
           (cons "point_count" (length pt_list))
         )
       )
       nil  ;; Not closed
     )
    )
    
    ;; Handle CIRCLE
    ((= ent_type "CIRCLE")
     (setq radius (cdr (assoc 40 ent_data)))
     (setq center (cdr (assoc 10 ent_data)))
     (setq area (* pi radius radius))
     (setq name (find-nearest-text center))
     (setq shape_type "circle")
     
     (list
       (cons "type" shape_type)
       (cons "name" (if name name "Unknown"))
       (cons "area" area)
       (cons "centroid" center)
       (cons "radius" radius)
       (cons "points" (list center))  ;; Just center point for circle
       (cons "point_count" 1)
     )
    )
    
    ;; Unknown type
    (t nil)
  )
)

;; Helper function: Get LWPOLYLINE points
(defun get-lwpolyline-points (ent / obj coords pt_list i count x y)
  (setq obj (vlax-ename->vla-object ent))
  (setq coords (vlax-safearray->list 
                (vlax-variant-value 
                  (vlax-get-property obj 'Coordinates))))
  
  (setq pt_list '())
  (setq i 0)
  (setq count (/ (length coords) 2))
  
  (repeat count
    (setq x (nth (* i 2) coords))
    (setq y (nth (+ (* i 2) 1) coords))
    (setq pt_list (append pt_list (list (list x y 0.0))))
    (setq i (1+ i))
  )
  
  pt_list
)

;; Helper function: Get POLYLINE points
(defun get-polyline-points (ent / pt_list vertex_ent vertex_data pt)
  (setq pt_list '())
  (setq vertex_ent (entnext ent))
  
  (while (and vertex_ent 
              (= (cdr (assoc 0 (entget vertex_ent))) "VERTEX"))
    (setq vertex_data (entget vertex_ent))
    (setq pt (cdr (assoc 10 vertex_data)))
    (setq pt_list (append pt_list (list pt)))
    (setq vertex_ent (entnext vertex_ent))
  )
  
  pt_list
)

;; Helper function: Save all shapes to JSON (with append support)
(defun save-all-shapes (shapes filename / file shape i j pt centroid old_content line old_shapes_count shape_id)
  ;; Check if file exists and read old content
  (setq old_content '())
  (setq old_shapes_count 0)
  
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
          
          ;; Count existing shapes by counting lines with "  {"
          (foreach line old_content
            (if (= line "  {")
              (setq old_shapes_count (1+ old_shapes_count))
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
      (if (= old_shapes_count 0)
        ;; First scan - create new array
        (progn
          (write-line "[" file)
        )
        ;; Existing shapes - copy old content without closing ]
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
      
      ;; Write new shapes
      (setq i 0)
      (foreach shape shapes
        ;; Generate unique ID for each shape
        (setq shape_id (generate-shape-id (cdr (assoc "name" shape)) (+ old_shapes_count i 1)))
        
        (write-line "  {" file)
        (write-line (strcat "    \"id\": \"" shape_id "\",") file)
        (write-line (strcat "    \"type\": \"" (cdr (assoc "type" shape)) "\",") file)
        (write-line (strcat "    \"name\": \"" (cdr (assoc "name" shape)) "\",") file)
        (write-line (strcat "    \"point_count\": " (itoa (cdr (assoc "point_count" shape))) ",") file)
        (write-line (strcat "    \"area\": " (rtos (cdr (assoc "area" shape)) 2 6) ",") file)
        
        ;; Centroid
        (setq centroid (cdr (assoc "centroid" shape)))
        (write-line (strcat "    \"centroid\": [" 
                           (rtos (car centroid) 2 6) ", " 
                           (rtos (cadr centroid) 2 6) 
                           "],") file)
        
        ;; Radius for circles
        (if (assoc "radius" shape)
          (write-line (strcat "    \"radius\": " (rtos (cdr (assoc "radius" shape)) 2 6) ",") file)
        )
        
        ;; Points
        (write-line "    \"points\": [" file)
        (setq j 0)
        (setq pt_list (cdr (assoc "points" shape)))
        (foreach pt pt_list
          (write-line (strcat "      [" 
                             (rtos (car pt) 2 6) ", " 
                             (rtos (cadr pt) 2 6) ", "
                             (rtos (if (caddr pt) (caddr pt) 0.0) 2 6) "]"
                             (if (< j (- (length pt_list) 1)) "," "")) file)
          (setq j (1+ j))
        )
        (write-line "    ]," file)
        
        ;; Timestamp and drawing
        (write-line (strcat "    \"timestamp\": \"" (rtos (getvar "CDATE") 2 0) "\",") file)
        (write-line (strcat "    \"drawing\": \"" (getvar "DWGNAME") "\"") file)
        
        ;; Close shape object
        (if (< i (- (length shapes) 1))
          (write-line "  }," file)
          (write-line "  }" file)
        )
        
        (setq i (1+ i))
      )
      
      ;; Close array
      (write-line "]" file)
      (close file)
      
      ;; Update total count
      (setq old_shapes_count (+ old_shapes_count (length shapes)))
      (princ (strcat "\n✓ Saved " (itoa (length shapes)) " new shape(s) to: " filename))
      (princ (strcat "\nTotal shapes: " (itoa old_shapes_count)))
    )
    (princ "\n✗ Error: Cannot create output file!")
  )
)

;; Auto-load message
(princ "\n=== CAD Column Inspector Pro Loaded ===")
(princ "\nType COLINSPECT to manually select polygon points")
(princ "\nType COLSCAN to auto-scan area for shapes")
(princ "\nType COLCLEAR to clear all saved polygons and start fresh")
(princ "\nType COLFINDPOLY to highlight a shape by ID")
(princ "\nType COLLIST to list all shape IDs in JSON file")
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

;; Helper function: Parse JSON file to find shape by ID
(defun find-shape-by-id (shape_id filename / file line content shapes shape found_shape shape_id_str)
  (setq found_shape nil)
  
  (if (findfile filename)
    (progn
      ;; Read entire file
      (setq file (open filename "r"))
      (setq content "")
      
      (if file
        (progn
          (while (setq line (read-line file))
            (setq content (strcat content line "\n"))
          )
          (close file)
          
          ;; Parse JSON
          (setq shapes (parse-json-shapes content))
          
          ;; Find shape with matching ID (case-insensitive comparison)
          (foreach shape shapes
            (setq shape_id_str (cdr (assoc "id" shape)))
            (if shape_id_str
              (progn
                ;; Compare IDs (case-insensitive and trimmed)
                (if (= (strcase (vl-string-trim " \t\n\r" shape_id) T)
                       (strcase (vl-string-trim " \t\n\r" shape_id_str) T))
                  (setq found_shape shape)
                )
              )
            )
          )
        )
      )
    )
  )
  
  found_shape
)

;; Helper function: Simple JSON parser for our shape format
(defun parse-json-shapes (content / shapes current_shape lines line in_shape in_points pt_list id_val name_val type_val area_val centroid_val radius_val)
  (setq shapes '())
  (setq current_shape nil)
  (setq in_shape nil)
  (setq in_points nil)
  (setq pt_list '())
  
  ;; Initialize shape properties
  (setq id_val nil)
  (setq name_val nil)
  (setq type_val nil)
  (setq area_val nil)
  (setq centroid_val nil)
  (setq radius_val nil)
  
  ;; Split into lines
  (setq lines (str-split content "\n"))
  
  (foreach line lines
    (cond
      ;; Start of shape - single { on line
      ((and (vl-string-search "{" line)
            (not (vl-string-search "\"" line)))
       (setq in_shape T)
       (setq in_points nil)
       (setq pt_list '())
       (setq id_val nil)
       (setq name_val nil)
       (setq type_val nil)
       (setq area_val nil)
       (setq centroid_val nil)
       (setq radius_val nil)
       (setq current_shape '())
      )
      
      ;; End of shape - single } on line
      ((and in_shape
            (vl-string-search "}" line)
            (not (vl-string-search "\"" line)))
       (if (and id_val name_val type_val)
         (progn
           (setq current_shape (list
             (cons "id" id_val)
             (cons "name" name_val)
             (cons "type" type_val)
             (cons "area" (if area_val area_val 0.0))
             (cons "centroid" (if centroid_val centroid_val '(0.0 0.0)))
             (cons "points" pt_list)
           ))
           ;; Add radius if it exists (for circles)
           (if radius_val
             (setq current_shape (append current_shape (list (cons "radius" radius_val))))
           )
           (setq shapes (append shapes (list current_shape)))
         )
       )
       (setq in_shape nil)
       (setq in_points nil)
      )
      
      ;; Inside shape - parse properties
      (in_shape
       (cond
         ;; Parse "id"
         ((vl-string-search "\"id\":" line)
          (setq id_val (extract-string-value line))
         )
         
         ;; Parse "name"
         ((vl-string-search "\"name\":" line)
          (setq name_val (extract-string-value line))
         )
         
         ;; Parse "type"
         ((vl-string-search "\"type\":" line)
          (setq type_val (extract-string-value line))
         )
         
         ;; Parse "area"
         ((vl-string-search "\"area\":" line)
          (setq area_val (extract-number-value line))
         )
         
         ;; Parse "centroid"
         ((vl-string-search "\"centroid\":" line)
          (setq centroid_val (parse-json-array-inline line))
         )
         
         ;; Parse "radius" (for circles)
         ((vl-string-search "\"radius\":" line)
          (setq radius_val (extract-number-value line))
         )
         
         ;; Start of points array
         ((vl-string-search "\"points\":" line)
          (setq in_points T)
         )
         
         ;; End of points array
         ((and in_points (vl-string-search "]," line) (not (vl-string-search "[" line)))
          (setq in_points nil)
         )
         
         ;; Parse point inside points array
         ((and in_points (vl-string-search "[" line))
          (setq pt_list (append pt_list (list (parse-json-point line))))
         )
       )
      )
    )
  )
  
  shapes
)

;; Helper function: Extract string value from JSON line
(defun extract-string-value (line / pos1 pos2 pos3 value_str)
  ;; Find the colon first
  (setq pos1 (vl-string-search ":" line))
  (if pos1
    (progn
      ;; Get substring after colon
      (setq line (substr line (+ pos1 2)))
      
      ;; Find first quote
      (setq pos2 (vl-string-search "\"" line))
      (if pos2
        (progn
          ;; Get substring after first quote
          (setq line (substr line (+ pos2 2)))
          
          ;; Find closing quote
          (setq pos3 (vl-string-search "\"" line))
          (if pos3
            (progn
              ;; Extract the value between quotes
              (setq value_str (substr line 1 pos3))
              value_str
            )
            nil
          )
        )
        nil
      )
    )
    nil
  )
)

;; Helper function: Extract number value from JSON line
(defun extract-number-value (line / start value_str)
  (setq start (vl-string-search ":" line))
  (if start
    (progn
      (setq value_str (substr line (+ start 2)))
      (setq value_str (vl-string-trim " ,\n\r\t" value_str))
      (atof value_str)
    )
    0.0
  )
)

;; Helper function: Parse inline JSON array like [1.0, 2.0]
(defun parse-json-array-inline (line / start end arr_str nums)
  (setq start (vl-string-search "[" line))
  (setq end (vl-string-search "]" line))
  (if (and start end)
    (progn
      (setq arr_str (substr line (+ start 2) (- end start 1)))
      (setq nums (str-split arr_str ","))
      (mapcar 'atof (mapcar '(lambda (s) (vl-string-trim " \t" s)) nums))
    )
    '(0.0 0.0)
  )
)

;; Helper function: Split string by delimiter
(defun str-split (str delim / pos result)
  (setq result '())
  (while (setq pos (vl-string-search delim str))
    (setq result (append result (list (substr str 1 pos))))
    (setq str (substr str (+ pos (strlen delim) 1)))
  )
  (if (> (strlen str) 0)
    (setq result (append result (list str)))
  )
  result
)

;; Helper function: Extract JSON key from line
(defun extract-json-key (line / start end)
  (if (setq start (vl-string-search "\"" line))
    (progn
      (setq line (substr line (+ start 2)))
      (if (setq end (vl-string-search "\"" line))
        (substr line 1 end)
        nil
      )
    )
    nil
  )
)

;; Helper function: Extract JSON value from line
(defun extract-json-value (line key / start value_str)
  (if (setq start (vl-string-search ": " line))
    (progn
      (setq value_str (substr line (+ start 3)))
      ;; Remove trailing comma and quotes
      (setq value_str (vl-string-trim " ,\n\r\t" value_str))
      (if (= (substr value_str 1 1) "\"")
        (progn
          ;; String value - remove quotes
          (setq value_str (substr value_str 2 (- (strlen value_str) 2)))
        )
        (progn
          ;; Number value
          (if (vl-string-search "[" value_str)
            ;; Array value (centroid)
            (parse-json-array value_str)
            ;; Simple number
            (atof value_str)
          )
        )
      )
    )
    nil
  )
)

;; Helper function: Parse JSON array [x, y]
(defun parse-json-array (str / nums)
  (setq str (vl-string-trim "[]" str))
  (setq nums (str-split str ", "))
  (mapcar 'atof nums)
)

;; Helper function: Parse JSON point [x, y, z]
(defun parse-json-point (line / str nums)
  (setq str (vl-string-trim " \n\r\t," line))
  (setq str (vl-string-trim "[]" str))
  (setq nums (str-split str ", "))
  (mapcar 'atof nums)
)

;; Helper function: Draw highlight around shape
;; Helper function: Pause for specified milliseconds
(defun pause-ms (milliseconds / start_time)
  (setq start_time (getvar "MILLISECS"))
  (while (< (- (getvar "MILLISECS") start_time) milliseconds)
    (grread t)  ; Non-blocking read, doesn't wait for input
  )
)

(defun highlight-shape (shape / shape_type points centroid radius pt1 pt2 i offset_points highlight_ent)
  (setq shape_type (cdr (assoc "type" shape)))
  (setq points (cdr (assoc "points" shape)))
  (setq centroid (cdr (assoc "centroid" shape)))
  
  (princ (strcat "\n✓ Shape found: " (cdr (assoc "name" shape))))
  (princ (strcat "\n  Type: " shape_type))
  (princ (strcat "\n  Area: " (rtos (cdr (assoc "area" shape)) 2 2) " sq units"))
  
  (cond
    ;; Circle
    ((= shape_type "circle")
     (setq radius (cdr (assoc "radius" shape)))
     (if (and radius centroid)
       (progn
         ;; Ensure lineweight display is on
         (setvar "LWDISPLAY" 1)
         
         ;; Draw yellow circle slightly larger
         (command "_.CIRCLE" 
                  (list (car centroid) (cadr centroid) 0.0)
                  (* radius 1.05))
         
         ;; Set to yellow and thicker lineweight
         (setq highlight_ent (entlast))
         (if highlight_ent
           (progn
             (command "_.CHPROP" highlight_ent "" "_C" "2" "_LW" "100" "")
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
    
    ;; Polygon/Polyline
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

;; Command: Find and highlight shape by ID
(defun C:COLFINDPOLY (/ shape_id result_file shape)
  (princ "\n=== Find Shape by ID ===")
  
  ;; Check if shape ID was set by script (for automation)
  (if (and (boundp '*COLFIND_SHAPEID*) *COLFIND_SHAPEID*)
    (progn
      (setq shape_id *COLFIND_SHAPEID*)
      (princ (strcat "\nSearching for: " shape_id))
      (setq *COLFIND_SHAPEID* nil)  ; Clear after use
    )
    ;; Otherwise, get shape ID from user (interactive mode)
    (setq shape_id (getstring "\nEnter Shape ID (e.g., COL_A1_001): "))
  )
  
  (if (and shape_id (> (strlen shape_id) 0))
    (progn
      (setq result_file (strcat (getvar "DWGPREFIX") "column_data.json"))
      
      (if (findfile result_file)
        (progn
          (princ "\nSearching for shape...")
          
          ;; Find the shape
          (setq shape (find-shape-by-id shape_id result_file))
          
          (if shape
            (progn
              (princ (strcat "\n✓ Found shape: " (cdr (assoc "name" shape))))
              (highlight-shape shape)
            )
            (progn
              (princ (strcat "\n✗ Shape not found: \"" shape_id "\""))
              (princ "\nTip: Use COLLIST to see all available shape IDs")
            )
          )
        )
        (progn
          (princ "\n✗ No shape data file found.")
          (princ "\nRun COLINSPECT or COLSCAN first to create data.")
        )
      )
    )
    (princ "\n✗ Invalid shape ID")
  )
  
  (princ)
)

;; Command: List all shape IDs in JSON file
(defun C:COLLIST (/ result_file file line content shapes shape count)
  (princ "\n=== List All Shape IDs ===")
  
  (setq result_file (strcat (getvar "DWGPREFIX") "column_data.json"))
  
  (if (findfile result_file)
    (progn
      (princ (strcat "\nReading: " result_file))
      
      ;; Read entire file
      (setq file (open result_file "r"))
      (setq content "")
      
      (if file
        (progn
          (while (setq line (read-line file))
            (setq content (strcat content line "\n"))
          )
          (close file)
          
          ;; Parse JSON
          (setq shapes (parse-json-shapes content))
          
          (if shapes
            (progn
              (princ (strcat "\n\n✓ Found " (itoa (length shapes)) " shape(s):\n"))
              (princ "----------------------------------------")
              
              (setq count 1)
              (foreach shape shapes
                (princ (strcat "\n" (itoa count) ". ID: " (cdr (assoc "id" shape))))
                (princ (strcat "\n   Name: " (cdr (assoc "name" shape))))
                (princ (strcat "\n   Type: " (cdr (assoc "type" shape))))
                (princ (strcat "\n   Area: " (rtos (cdr (assoc "area" shape)) 2 2) " sq units"))
                (princ "\n")
                (setq count (1+ count))
              )
              
              (princ "\nUse COLFINDPOLY <ID> to highlight a shape")
            )
            (princ "\n✗ No shapes found in JSON file")
          )
        )
        (princ "\n✗ Error reading JSON file")
      )
    )
    (progn
      (princ "\n✗ No shape data file found.")
      (princ (strcat "\nExpected location: " result_file))
      (princ "\nRun COLINSPECT or COLSCAN first to create data.")
    )
  )
  
  (princ)
)
