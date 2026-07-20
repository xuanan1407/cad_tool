;; commands.lsp
;; Main command definitions
;; Copyright (c) 2026 Tran Xuan An

;; ========================================
;; COMMAND: COLINSPECT
;; Manual polygon point selection
;; ========================================

(defun C:COLINSPECT (/ pt_list pt count area centroid result_file confirm redline_list 
                      nearest_text current_shape history_shapes global_shapes error_count
                      prev_pt last_ent ent_data)
  (print-header "CAD Column Inspector Pro")
  (princ "\nSelect points to create polygon (Press Enter to finish)")
  
  ;; Initialize
  (setq pt_list '())
  (setq count 0)
  (setq redline_list '())
  
  ;; Get points from user
  (while (setq pt (if (null pt_list)
                    (getpoint "\nSelect first point: ")
                    (getpoint (car (reverse pt_list)) "\nSelect next point (or press Enter to finish): ")))
    (setq pt_list (append pt_list (list pt)))
    (setq count (1+ count))
    (princ (strcat "\nPoint " (itoa count) " selected: " (vl-princ-to-string pt)))
    
    ;; Draw red line from previous point to current point (ONE segment at a time)
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
  
  ;; Check minimum points
  (if (< (length pt_list) 3)
    (progn
      (print-error "Need at least 3 points to create a polygon!")
      (delete-entities redline_list)
    )
    (progn
      ;; Delete preview lines
      (delete-entities redline_list)
      
      ;; Draw final polygon
      (draw-final-polygon pt_list)
      
      ;; Calculate properties
      (setq area (calculate-polygon-area pt_list))
      (setq centroid (calculate-centroid pt_list))
      (setq nearest_text (find-nearest-text centroid))
      (if (null nearest_text) (setq nearest_text "[No text found]"))
      
      ;; Load history and check consistency
      (setq result_file (get-result-file))
      (setq history_shapes '())
      
      (if (file-exists-p result_file)
        (setq history_shapes (load-shapes-from-json result_file))
      )
      
      ;; Create current shape data
      (setq current_shape (list (cons "name" nearest_text) (cons "area" area)))
      
      ;; Combine and check
      (setq global_shapes (append history_shapes (list current_shape)))
      (setq error_count (check-area-consistency global_shapes))
      
      ;; Display info
      (princ (strcat "\n\nPolygon created with " (itoa count) " points"))
      (princ (strcat "\nName: " nearest_text))
      (princ (strcat "\nArea: " (rtos area 2 2) " sq units"))
      (princ (strcat "\nCentroid: " (vl-princ-to-string centroid)))
      
      ;; Ask to confirm
      (initget "Yes No")
      (if (= error_count 0)
        (setq confirm (getkword "\nExport this polygon data? [Yes/No] <Yes>: "))
        (setq confirm (getkword "\n⚠ WARNING: Data contains inconsistencies! Force export? [Yes/No] <No>: "))
      )
      
      (if (or (and (null confirm) (= error_count 0)) (= confirm "Yes"))
        (progn
          ;; Save data
          (append-polygon-data pt_list area centroid nearest_text result_file)
          
          ;; Delete temporary polyline
          (command "_.ERASE" "L" "")
          
          ;; Call Python
          (call-python-processor result_file)
        )
        (progn
          ;; Cancel - delete polyline
          (command "_.ERASE" "L" "")
          (princ "\nOperation cancelled.")
        )
      )
    )
  )
  (princ)
)

;; ========================================
;; COMMAND: COLSCAN
;; Auto-scan area for shapes
;; ========================================

(defun C:COLSCAN (/ p1 p2 ss i ent shape_data all_shapes result_file 
                   original_count duplicate_count confirm tolerance_input tolerance
                   history_shapes global_shapes error_count)
  (print-header "CAD Column Auto-Scanner")
  
  ;; Ask for tolerance
  (setq tolerance_input (getreal (strcat "\nEnter minimum distance between shapes (default " 
                                        (rtos *COL_DEFAULT_TOLERANCE* 2 1) "mm): ")))
  (if (null tolerance_input)
    (setq tolerance *COL_DEFAULT_TOLERANCE*)
    (setq tolerance tolerance_input)
  )
  
  (print-success (strcat "Tolerance: " (rtos tolerance 2 2) "mm (centroid to centroid)"))
  (princ "\n✓ Shapes closer than this will be removed (keeping first one)")
  (princ "\nSelect area to scan for shapes...")
  
  ;; Get selection window
  (setq p1 (getpoint "\nSelect first corner: "))
  
  (if p1
    (progn
      (setq p2 (getcorner p1 "\nSelect opposite corner: "))
      
      (if p2
        (progn
          ;; Select all shapes
          (setq ss (ssget "_W" p1 p2 *COL_SCAN_FILTER*))
          
          (if ss
            (progn
              (setq all_shapes '())
              (setq i 0)
              
              (print-success (strcat "Found " (itoa (sslength ss)) " shape(s) in selected area"))
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
              
              ;; Remove duplicates
              (setq original_count (length all_shapes))
              (setq all_shapes (remove-duplicate-shapes all_shapes tolerance))
              (setq duplicate_count (- original_count (length all_shapes)))
              
              ;; Load history and check consistency
              (setq result_file (get-result-file))
              (setq history_shapes '())
              
              (if (file-exists-p result_file)
                (setq history_shapes (load-shapes-from-json result_file))
              )
              
              ;; Combine and check
              (setq global_shapes (append history_shapes all_shapes))
              (setq error_count (check-area-consistency global_shapes))
              
              ;; Display summary
              (princ (strcat "\n\n✓ Total shapes detected: " (itoa original_count)))
              (if (> duplicate_count 0)
                (print-warning (strcat "Removed " (itoa duplicate_count) " shape(s) within " 
                                      (rtos tolerance 2 2) "mm"))
              )
              (print-success (strcat "Unique shapes: " (itoa (length all_shapes))))
              
              ;; Ask to save
              (if (> (length all_shapes) 0)
                (progn
                  (initget "Yes No")
                  (setq confirm (getkword "\nExport all shapes to JSON? [Yes/No] <Yes>: "))
                  
                  (if (or (null confirm) (= confirm "Yes"))
                    (progn
                      ;; Save all shapes
                      (save-all-shapes all_shapes result_file)
                      
                      ;; Call Python
                      (call-python-processor result_file)
                    )
                    (princ "\nOperation cancelled.")
                  )
                )
                (print-warning "No valid shapes found!")
              )
            )
            (print-warning "No shapes found in selected area!")
          )
        )
        (princ "\nSelection cancelled.")
      )
    )
    (princ "\nSelection cancelled.")
  )
  (princ)
)

;; ========================================
;; COMMAND: COLCLEAR
;; Clear all saved polygon data
;; ========================================

(defun C:COLCLEAR (/ result_file confirm)
  (print-header "Clear Polygon Data")
  
  ;; Ask confirmation
  (initget "Yes No")
  (setq confirm (getkword "\nDelete all saved polygons and start fresh? [Yes/No] <No>: "))
  
  (if (= confirm "Yes")
    (progn
      (setq result_file (get-result-file))
      
      (if (file-exists-p result_file)
        (progn
          (vl-file-delete result_file)
          (print-success "All polygon data cleared!")
          (princ "\nYou can now start selecting new polygons with COLINSPECT")
        )
        (print-warning "No polygon data file found.")
      )
    )
    (princ "\nOperation cancelled.")
  )
  (princ)
)

;; ========================================
;; COMMAND: COLFINDPOLY
;; Find and highlight shape by ID (reverse lookup)
;; ========================================

(defun C:COLFINDPOLY (/ shape_id result_file shape)
  (print-header "Find Shape by ID")
  
  ;; Check if shape ID is passed via global variable (from Python/Excel)
  (if (and (boundp '*COLFIND_SHAPEID*) *COLFIND_SHAPEID*)
    (progn
      (setq shape_id *COLFIND_SHAPEID*)
      (princ (strcat "\nSearching for: " shape_id))
      (setq *COLFIND_SHAPEID* nil)  ; Clear after use
    )
    ;; Otherwise, ask user
    (setq shape_id (getstring T "\nEnter shape ID (e.g., C1_001): "))
  )
  
  (if shape_id
    (progn
      (setq result_file (get-result-file))
      
      (if (file-exists-p result_file)
        (progn
          (princ (strcat "\nSearching for shape: " shape_id "..."))
          (setq shape (find-shape-by-id shape_id result_file))
          
          (if shape
            ;; Use highlight-shape function (handles circle and polygon correctly)
            (highlight-shape shape)
            (print-error (strcat "Shape ID not found: " shape_id))
          )
        )
        (print-error "No polygon data file found. Run COLINSPECT or COLSCAN first.")
      )
    )
    (print-error "No shape ID provided!")
  )
  
  (princ)
)

;; ========================================
;; COMMAND: COLLIST
;; List all shape IDs in JSON file
;; ========================================

(defun C:COLLIST (/ result_file file line count)
  (print-header "List All Shape IDs")
  
  (setq result_file (get-result-file))
  
  (if (file-exists-p result_file)
    (progn
      (setq file (open result_file "r"))
      (setq count 0)
      
      (if file
        (progn
          (princ "\nShape IDs in database:")
          (princ (strcat "\n" (vl-string-repeat "-" 40)))
          
          (while (setq line (read-line file))
            (if (vl-string-search "\"id\":" line)
              (progn
                (setq count (1+ count))
                (princ (strcat "\n" (itoa count) ". " (vl-string-trim " \t,\"" 
                              (substr line (+ (vl-string-search ":" line) 2)))))
              )
            )
          )
          
          (close file)
          (princ (strcat "\n" (vl-string-repeat "-" 40)))
          (print-success (strcat "Total: " (itoa count) " shapes"))
        )
        (print-error "Cannot open data file!")
      )
    )
    (print-warning "No polygon data file found.")
  )
  (princ)
)

(princ "\n✓ Commands loaded")
(princ)
