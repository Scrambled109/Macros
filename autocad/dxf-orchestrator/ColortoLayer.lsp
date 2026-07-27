(defun c:ColorToLayer ( / colorMap mapItem colorIndex targetLayer ss i entData )
  (vl-load-com)

  ;; 1. PROCESS EXPLICIT COLORS
  ;; Format: (AutoCAD_Color_Index . "Exact_Layer_Name")
  (setq colorMap '(
    (1 . "CUT - OUTSIDE STRAIGHT")  ;; Red
    (2 . "PLOT")                    ;; Yellow
    (3 . "PIN STAMP LINE MARKING")  ;; Green
    (5 . "CUT - INSIDE STRAIGHT")   ;; Blue
    (6 . "PIN STAMP TEXT")          ;; Magenta
    (7 . "PIN STAMP TEXT")          ;; White/Black (Explicitly tagged)
  ))

  ;; Loop through explicit color mapping
  (foreach mapItem colorMap
    (setq colorIndex (car mapItem))
    (setq targetLayer (cdr mapItem))

    (if (setq ss (ssget "X" (list (cons 62 colorIndex))))
      (progn
        (setq i 0)
        (while (< i (sslength ss))
          (setq entData (entget (ssname ss i)))
          
          ;; Change layer
          (setq entData (subst (cons 8 targetLayer) (assoc 8 entData) entData))
          ;; Change explicit color to 'ByLayer' (256)
          (setq entData (subst (cons 62 256) (assoc 62 entData) entData))
          
          (entmod entData)
          (setq i (1+ i))
        )
      )
    )
  )

  ;; 2. SWEEP DEFAULT LAYER 0 ITEMS (The "ByLayer" White lines)
  ;; Select all objects currently sitting on Layer 0
  (if (setq ss (ssget "X" '((8 . "0"))))
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq entData (entget (ssname ss i)))
        
        ;; If the entity is missing the DXF 62 code, its color is "ByLayer".
        ;; On Layer 0, "ByLayer" visually means White.
        (if (not (assoc 62 entData))
          (progn
            ;; Move it to PIN STAMP TEXT
            (setq entData (subst (cons 8 "PIN STAMP TEXT") (assoc 8 entData) entData))
            (entmod entData)
          )
        )
        (setq i (1+ i))
      )
    )
  )

  (princ "\nLine processing complete! Objects moved to mapped layers.")
  (princ)
)