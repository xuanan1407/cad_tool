;; geometry.lsp
;; Geometric calculation functions
;; Copyright (c) 2026 Tran Xuan An

;; ========================================
;; AREA CALCULATIONS
;; ========================================

(defun calculate-polygon-area (pt_list / area i n x1 y1 x2 y2)
  "Calculate polygon area using Shoelace formula"
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

;; ========================================
;; CENTROID CALCULATIONS
;; ========================================

(defun calculate-centroid (pt_list / sum_x sum_y count pt)
  "Calculate centroid (geometric center) of points"
  (setq sum_x 0.0)
  (setq sum_y 0.0)
  (setq count (length pt_list))
  
  (foreach pt pt_list
    (setq sum_x (+ sum_x (car pt)))
    (setq sum_y (+ sum_y (cadr pt)))
  )
  
  (list (/ sum_x count) (/ sum_y count))
)

;; ========================================
;; DISTANCE CALCULATIONS
;; ========================================

(defun distance-2d (pt1 pt2)
  "Calculate 2D distance between two points"
  (sqrt (+ 
    (* (- (car pt1) (car pt2)) (- (car pt1) (car pt2)))
    (* (- (cadr pt1) (cadr pt2)) (- (cadr pt1) (cadr pt2)))
  ))
)

;; ========================================
;; POLYLINE UTILITIES
;; ========================================

(defun get-lwpolyline-points (ent / obj coords pt_list i count x y)
  "Extract point list from LWPOLYLINE"
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

(defun get-polyline-points (ent / pt_list vertex_ent vertex_data pt)
  "Extract point list from POLYLINE (old-style)"
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

;; ========================================
;; DUPLICATE DETECTION
;; ========================================

(defun is-duplicate-shape (shape1 shape2 tolerance / centroid1 centroid2 dist)
  "Check if two shapes are too close (within tolerance)"
  (setq centroid1 (cdr (assoc "centroid" shape1)))
  (setq centroid2 (cdr (assoc "centroid" shape2)))
  
  (if (and centroid1 centroid2)
    (progn
      (setq dist (distance-2d centroid1 centroid2))
      (< dist tolerance)
    )
    nil
  )
)

(defun remove-duplicate-shapes (shapes tolerance / unique_shapes shape is_dup unique_shape)
  "Remove shapes that are too close to each other (keep first)"
  (setq unique_shapes '())
  
  (foreach shape shapes
    (setq is_dup nil)
    
    ;; Check if this shape is too close to any unique shape
    (foreach unique_shape unique_shapes
      (if (is-duplicate-shape shape unique_shape tolerance)
        (setq is_dup T)
      )
    )
    
    ;; Add if not duplicate
    (if (not is_dup)
      (setq unique_shapes (append unique_shapes (list shape)))
    )
  )
  
  unique_shapes
)

;; ========================================
;; CONSISTENCY CHECKS
;; ========================================

(defun check-area-consistency (shapes / checklist errors item name area existing_area)
  "Check for area inconsistencies within same shape name"
  (setq checklist '())
  (setq errors 0)
  
  (princ "\n--- CHECKING SHAPE AREA CONSISTENCY ---")
  
  (foreach item shapes
    (setq name (cdr (assoc "name" item)))
    (setq area (cdr (assoc "area" item)))
    
    ;; Check if this name has been seen before
    (setq existing_area (cdr (assoc name checklist)))
    
    (if existing_area
      ;; Name exists - check if area matches
      (if (> (abs (- area existing_area)) *COL_AREA_TOLERANCE*)
        (progn
          (print-warning (strcat "Inconsistent area for [" name "]: " 
                         (rtos existing_area 2 2) " vs " (rtos area 2 2)))
          (setq errors (1+ errors))
        )
      )
      ;; New name - store as baseline
      (setq checklist (append checklist (list (cons name area))))
    )
  )
  
  (if (> errors 0)
    (princ (strcat "\n\n❌ ERROR: Found " (itoa errors) " geometric inconsistency/inconsistencies!"))
    (princ "\n✓ All shapes with same name have consistent areas")
  )
  
  errors
)

(princ "\n✓ Geometry functions loaded")
(princ)
