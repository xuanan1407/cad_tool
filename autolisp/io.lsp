;; io.lsp
;; File I/O functions for JSON export/import
;; Copyright (c) 2026 Tran Xuan An

;; ========================================
;; JSON EXPORT (SINGLE SHAPE)
;; ========================================

(defun append-polygon-data (pt_list area centroid name filename / old_content lines polygon_count file i j pt shape_id)
  "Append single polygon data to JSON array"
  
  ;; Read existing file
  (setq old_content '())
  (setq polygon_count 0)
  
  (if (findfile filename)
    (progn
      (setq file (open filename "r"))
      (if file
        (progn
          (while (setq line (read-line file))
            (setq old_content (append old_content (list line)))
          )
          (close file)
          
          ;; Count existing polygons
          (foreach line old_content
            (if (= line "  {")
              (setq polygon_count (1+ polygon_count))
            )
          )
        )
      )
    )
  )
  
  ;; Write new file
  (setq file (open filename "w"))
  
  (if file
    (progn
      (if (= polygon_count 0)
        ;; First polygon - create array
        (write-line "[" file)
        ;; Append to existing - copy old content without closing ]
        (progn
          (setq i 0)
          (foreach line old_content
            (if (< i (- (length old_content) 1))
              (progn
                ;; Add comma to last shape
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
      
      ;; Write new polygon
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
      
      ;; Metadata
      (write-line (strcat "    \"timestamp\": \"" (rtos (getvar "CDATE") 2 0) "\",") file)
      (write-line (strcat "    \"drawing\": \"" (getvar "DWGNAME") "\"") file)
      
      ;; Close shape
      (write-line "  }" file)
      
      ;; Close array
      (write-line "]" file)
      
      (close file)
      
      ;; Report
      (setq polygon_count (1+ polygon_count))
      (print-success (strcat "Polygon #" (itoa polygon_count) " saved to: " filename))
      (princ (strcat "\nTotal polygons: " (itoa polygon_count)))
    )
    (print-error "Cannot create output file!")
  )
)

;; ========================================
;; JSON EXPORT (MULTIPLE SHAPES)
;; ========================================

(defun save-all-shapes (shapes filename / file shape i j pt centroid old_content line old_shapes_count shape_id pt_list)
  "Save all shapes to JSON array (with append support)"
  
  ;; Read existing file
  (setq old_content '())
  (setq old_shapes_count 0)
  
  (if (findfile filename)
    (progn
      (setq file (open filename "r"))
      (if file
        (progn
          (while (setq line (read-line file))
            (setq old_content (append old_content (list line)))
          )
          (close file)
          
          ;; Count existing shapes
          (foreach line old_content
            (if (= line "  {")
              (setq old_shapes_count (1+ old_shapes_count))
            )
          )
        )
      )
    )
  )
  
  ;; Write new file
  (setq file (open filename "w"))
  
  (if file
    (progn
      (if (= old_shapes_count 0)
        ;; First scan - create array
        (write-line "[" file)
        ;; Append to existing
        (progn
          (setq i 0)
          (foreach line old_content
            (if (< i (- (length old_content) 1))
              (progn
                ;; Add comma to last shape
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
        ;; Generate unique ID
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
        
        ;; Metadata
        (write-line (strcat "    \"timestamp\": \"" (rtos (getvar "CDATE") 2 0) "\",") file)
        (write-line (strcat "    \"drawing\": \"" (getvar "DWGNAME") "\"") file)
        
        ;; Close shape
        (if (< i (- (length shapes) 1))
          (write-line "  }," file)
          (write-line "  }" file)
        )
        
        (setq i (1+ i))
      )
      
      ;; Close array
      (write-line "]" file)
      (close file)
      
      ;; Report
      (setq old_shapes_count (+ old_shapes_count (length shapes)))
      (print-success (strcat "Saved " (itoa (length shapes)) " new shape(s) to: " filename))
      (princ (strcat "\nTotal shapes: " (itoa old_shapes_count)))
    )
    (print-error "Cannot create output file!")
  )
)

;; ========================================
;; JSON IMPORT
;; ========================================

(defun load-shapes-from-json (file_path / f_in line shapes current_name current_area pos_colon raw_val)
  "Load shapes from JSON file (simplified parser)"
  (setq shapes '())
  (setq f_in (open file_path "r"))
  
  (if f_in
    (progn
      (while (setq line (read-line f_in))
        (setq line (vl-string-trim " \t," line))
        
        ;; Extract "name"
        (if (vl-string-search "\"name\":" line)
          (progn
            (setq pos_colon (vl-string-search ":" line))
            (setq current_name (vl-string-trim " \"\t," (substr line (+ pos_colon 2))))
          )
        )
        
        ;; Extract "area"
        (if (vl-string-search "\"area\":" line)
          (progn
            (setq pos_colon (vl-string-search ":" line))
            (setq raw_val (vl-string-trim " \"\t," (substr line (+ pos_colon 2))))
            (setq current_area (distof raw_val))
            
            ;; Once both extracted, add to list
            (if (and current_name current_area)
              (progn
                (setq shapes (append shapes (list (list (cons "name" current_name) (cons "area" current_area)))))
                (setq current_name nil)
                (setq current_area nil)
              )
            )
          )
        )
      )
      (close f_in)
    )
  )
  
  shapes
)

;; ========================================
;; FIND SHAPE BY ID
;; ========================================

(defun find-shape-by-id (shape_id filename / file line content shapes shape found_shape shape_id_str)
  "Find shape by ID in JSON file"
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
          
          ;; Find matching ID (case-insensitive)
          (foreach shape shapes
            (setq shape_id_str (cdr (assoc "id" shape)))
            (if shape_id_str
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
  
  found_shape
)

;; ========================================
;; FIND SHAPES BY PREFIX (GROUP SEARCH)
;; ========================================

(defun find-shapes-by-prefix (prefix filename / file line content shapes shape found_shapes shape_id_str)
  "Find all shapes with IDs starting with the given prefix"
  (setq found_shapes '())
  
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
          
          ;; Find all shapes with matching prefix (case-insensitive)
          (foreach shape shapes
            (setq shape_id_str (cdr (assoc "id" shape)))
            (if shape_id_str
              (progn
                (setq shape_id_str (vl-string-trim " \t\n\r" shape_id_str))
                ;; Check if ID starts with prefix
                (if (= (strcase (substr shape_id_str 1 (strlen prefix)) T)
                       (strcase prefix T))
                  (setq found_shapes (append found_shapes (list shape)))
                )
              )
            )
          )
        )
      )
    )
  )
  
  found_shapes
)

;; ========================================
;; JSON PARSER (SIMPLIFIED)
;; ========================================

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

;; ========================================
;; JSON HELPER FUNCTIONS
;; ========================================

(defun extract-string-value (line / pos1 pos2 pos3 value_str)
  "Extract string value from JSON line"
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

(defun extract-number-value (line / start value_str)
  "Extract number value from JSON line"
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

(defun parse-json-array-inline (line / start end arr_str nums)
  "Parse inline JSON array like [1.0, 2.0]"
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

(defun parse-json-point (line / str nums)
  "Parse JSON point [x, y, z]"
  (setq str (vl-string-trim " \n\r\t," line))
  (setq str (vl-string-trim "[]" str))
  (setq nums (str-split str ", "))
  (mapcar 'atof nums)
)

(princ "\n✓ I/O functions loaded")
(princ)
