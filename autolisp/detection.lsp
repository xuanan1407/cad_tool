;; detection.lsp
;; Shape detection and analysis functions
;; Copyright (c) 2026 Tran Xuan An

;; ========================================
;; ENTITY ANALYSIS
;; ========================================

(defun analyze-entity (ent / ent_data ent_type obj area centroid pt_list name shape_type radius center)
  "Analyze entity and extract shape data"
  (setq ent_data (entget ent))
  (setq ent_type (cdr (assoc 0 ent_data)))
  (setq obj (vlax-ename->vla-object ent))
  
  (cond
    ;; ========================================
    ;; LWPOLYLINE (lightweight polyline)
    ;; ========================================
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
       nil  ; Not closed - skip
     )
    )
    
    ;; ========================================
    ;; POLYLINE (old-style polyline)
    ;; ========================================
    ((= ent_type "POLYLINE")
     (if (= (logand (cdr (assoc 70 ent_data)) 1) 1)  ; Check closed flag
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
       nil  ; Not closed - skip
     )
    )
    
    ;; ========================================
    ;; CIRCLE
    ;; ========================================
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
       (cons "points" (list center))
       (cons "point_count" 1)
     )
    )
    
    ;; Unknown type - skip
    (t nil)
  )
)

;; ========================================
;; VISUAL PREVIEW
;; ========================================

(defun draw-preview-polyline (pt_list / prev_pt current_pt redline_list last_ent ent_data i)
  "Draw red preview polylines between points"
  (setq redline_list '())
  
  (if (> (length pt_list) 1)
    (progn
      (setq i 1)
      (while (< i (length pt_list))
        (setq prev_pt (nth (- i 1) pt_list))
        (setq current_pt (nth i pt_list))
        
        ;; Draw thick red polyline
        (command "_.PLINE" prev_pt "W" (rtos *COL_PREVIEW_WIDTH*) (rtos *COL_PREVIEW_WIDTH*) current_pt "")
        
        ;; Change color to red
        (setq last_ent (entlast))
        (if last_ent
          (progn
            (setq ent_data (entget last_ent))
            
            ;; Set color
            (if (assoc 62 ent_data)
              (setq ent_data (subst (cons 62 *COL_PREVIEW_COLOR*) (assoc 62 ent_data) ent_data))
              (setq ent_data (append ent_data (list (cons 62 *COL_PREVIEW_COLOR*))))
            )
            
            (entmod ent_data)
            (entupd last_ent)
            
            (setq redline_list (append redline_list (list last_ent)))
          )
        )
        
        (setq i (1+ i))
      )
    )
  )
  
  redline_list
)

(defun draw-final-polygon (pt_list / pt)
  "Draw final closed polyline"
  (command "_.PLINE")
  (foreach pt pt_list
    (command pt)
  )
  (command "_C")  ; Close
)

(princ "\n✓ Detection functions loaded")
(princ)
