(defun c:ColorToLayer ( / colorMap mapItem colorIndex targetLayer ss i entData )
  (vl-load-com)

  ;; Define your color-to-layer mappings here.
  ;; Format: (AutoCAD_Color_Index . "Exact_Layer_Name")
  (setq colorMap '(
    (1 . "CUT - OUTSIDE STRAIGHT")  ;; Red
    (2 . "PLOT")                    ;; Yellow
    (3 . "PIN STAMP LINE MARKING")  ;; Green
    (5 . "CUT - INSIDE STRAIGHT")   ;; Blue
    (6 . "PIN STAMP TEXT")          ;; Magenta (Overriding Scribe per your request)
    (7 . "PIN STAMP TEXT")          ;; White/Black
    ;; Add any other unique color mappings below:
    ;; (Color_Number . "Layer_Name")
  ))

  ;; Loop through each mapping defined above
  (foreach mapItem colorMap
    (setq colorIndex (car mapItem))
    (setq targetLayer (cdr mapItem))

    ;; Select all objects in the drawing that have this specific color
    (if (setq ss (ssget "X" (list (cons 62 colorIndex))))
      (progn
        (setq i 0)
        (while (< i (sslength ss))
          (setq entData (entget (ssname ss i)))
          
          ;; 1. Change the object's layer (DXF code 8)
          (setq entData (subst (cons 8 targetLayer) (assoc 8 entData) entData))
          
          ;; 2. Change the object's explicit color to 'ByLayer' (DXF code 62 -> 256)
          ;; This ensures it visually takes on the properties of its new layer
          (setq entData (subst (cons 62 256) (assoc 62 entData) entData))
          
          ;; Apply the changes to the entity
          (entmod entData)
          (setq i (1+ i))
        )
      )
    )
  )
  (princ "\nLine processing complete! Objects moved to mapped layers.")
  (princ)
)