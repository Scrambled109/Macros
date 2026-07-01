(vl-load-com)

(defun c:LayerToggle ( / ss i ent edata minpt maxpt pt layerName txtHeight lblLayer)
  (setq lblLayer "Z_TEMP_LAYER_LABELS")
  
  ;; Check if the temporary labels already exist
  (if (setq ss (ssget "X" (list (cons 8 lblLayer))))
    (progn
      ;; If they exist, delete them (Toggle OFF)
      (command "_.ERASE" ss "")
      (princ "\nLayer labels cleared.")
    )
    (progn
      ;; If they don't exist, generate them (Toggle ON)
      ;; Create a non-printing, yellow layer for the labels
      (command "_.LAYER" "_M" lblLayer "_C" "2" "" "_P" "N" "" "") 
      
      ;; Dynamically size the text so it's readable based on your current zoom level
      (setq txtHeight (* (getvar "VIEWSIZE") 0.012))
      
      ;; Grab all lines, polylines, arcs, and circles
      (if (setq ss (ssget "X" '((0 . "LINE,LWPOLYLINE,CIRCLE,ARC"))))
        (progn
          (setq i 0)
          (while (< i (sslength ss))
            (setq ent (ssname ss i))
            (setq edata (entget ent))
            (setq layerName (cdr (assoc 8 edata)))
            
            ;; Find the center point of the object
            (vla-getboundingbox (vlax-ename->vla-object ent) 'minpt 'maxpt)
            (setq pt (mapcar '/ (mapcar '+ (vlax-safearray->list minpt) (vlax-safearray->list maxpt)) '(2.0 2.0 2.0)))
            
            ;; Generate the text stamp
            (entmake
              (list
                '(0 . "TEXT")
                (cons 8 lblLayer)
                (cons 10 pt)
                (cons 40 txtHeight)
                (cons 1 layerName)
                '(50 . 0.0) 
                '(72 . 1) ; Center justified
                (cons 11 pt)
              )
            )
            (setq i (1+ i))
          )
          (princ (strcat "\nX-Ray active: Stamped " (itoa (sslength ss)) " objects."))
        )
        (princ "\nNo valid lines found to label.")
      )
    )
  )
  (princ)
)