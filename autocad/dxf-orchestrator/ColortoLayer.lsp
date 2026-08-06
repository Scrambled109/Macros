(vl-load-com)

(defun SPC:EntityText (entData / combined pair)
  (setq combined "")
  (foreach pair entData
    (if (or (= (car pair) 1) (= (car pair) 3))
      (setq combined (strcat combined " " (cdr pair)))
    )
  )
  combined
)

(defun SPC:FallbackBevelTextP (value / upper)
  (setq upper (strcase value))
  (or
    (vl-string-search "BEVEL" upper)
    (vl-string-search "BVL" upper)
    (vl-string-search "CHAMFER" upper)
    (vl-string-search "SNIPE" upper)
    (vl-string-search "GOUGE" upper)
    (vl-string-search "->" upper)
    (vl-string-search "<-" upper)
    (vl-string-search "←" value)
    (vl-string-search "↑" value)
    (vl-string-search "→" value)
    (vl-string-search "↓" value)
  )
)

(defun SPC:BevelTextP (value / regex match result)
  (setq result nil)
  (if (and value (/= value ""))
    (progn
      (setq regex
        (vl-catch-all-apply 'vlax-create-object (list "VBScript.RegExp"))
      )
      (if (vl-catch-all-error-p regex)
        (setq result (SPC:FallbackBevelTextP value))
        (progn
          (vlax-put-property regex 'Global :vlax-false)
          (vlax-put-property regex 'IgnoreCase :vlax-true)
          (vlax-put-property regex 'Pattern
            "\\b(BEVEL(ED)?|BVL|CHAMFER|SNIP(E|ED|ING)?|BACK[ \\t]+GOUGE|GOUGE)\\b|\\b(K|V|RV)\\b|\\b(K|V|RV)[ \\t]*[-–—:/]?[ \\t]*[0-9]+(\\.[0-9]+)?\\b|(-{1,2}>|<{1,2}-|←|↑|→|↓|↖|↗|↘|↙|⇐|⇑|⇒|⇓|➔|➜|➤|►|◄|△|▽)"
          )
          (setq match
            (vl-catch-all-apply 'vlax-invoke-method
              (list regex 'Test value)
            )
          )
          (if (vl-catch-all-error-p match)
            (setq result (SPC:FallbackBevelTextP value))
            (setq result (= match :vlax-true))
          )
          (vlax-release-object regex)
        )
      )
    )
  )
  result
)

(defun SPC:MoveToPlot (entity / entData)
  (setq entData (entget entity))
  (setq entData (subst (cons 8 "PLOT") (assoc 8 entData) entData))
  (if (assoc 62 entData)
    (setq entData (subst (cons 62 256) (assoc 62 entData) entData))
  )
  (entmod entData)
  (entupd entity)
)

(defun c:ColorToLayer
  ( / colorMap mapItem colorIndex targetLayer ss i entity entData )

  ;; First apply the standard explicit-color mapping.
  (setq colorMap '(
    (1 . "CUT - OUTSIDE STRAIGHT")
    (2 . "PLOT")
    (3 . "PIN STAMP LINE MARKING")
    (5 . "CUT - INSIDE STRAIGHT")
    (6 . "PIN STAMP TEXT")
    (7 . "PIN STAMP TEXT")
    (8 . "PLOT")
  ))

  (foreach mapItem colorMap
    (setq colorIndex (car mapItem))
    (setq targetLayer (cdr mapItem))
    (if (setq ss (ssget "X" (list (cons 62 colorIndex))))
      (progn
        (setq i 0)
        (while (< i (sslength ss))
          (setq entity (ssname ss i))
          (setq entData (entget entity))
          (setq entData (subst (cons 8 targetLayer) (assoc 8 entData) entData))
          (setq entData (subst (cons 62 256) (assoc 62 entData) entData))
          (entmod entData)
          (setq i (1+ i))
        )
      )
    )
  )

  ;; ByLayer white entities left on layer 0 are ordinary pin-stamp text.
  (if (setq ss (ssget "X" '((8 . "0"))))
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq entity (ssname ss i))
        (setq entData (entget entity))
        (if (not (assoc 62 entData))
          (progn
            (setq entData
              (subst (cons 8 "PIN STAMP TEXT") (assoc 8 entData) entData)
            )
            (entmod entData)
          )
        )
        (setq i (1+ i))
      )
    )
  )

  ;; Override the color mapping for bevel annotations. Bevel TEXT/MTEXT and
  ;; every leader/multileader arrow belong on PLOT, never on a pin-stamp layer.
  (if (setq ss (ssget "X" '((0 . "LEADER,MLEADER,MULTILEADER"))))
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (SPC:MoveToPlot (ssname ss i))
        (setq i (1+ i))
      )
    )
  )

  (if (setq ss (ssget "X" '((0 . "TEXT,MTEXT"))))
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq entity (ssname ss i))
        (setq entData (entget entity))
        (if (SPC:BevelTextP (SPC:EntityText entData))
          (SPC:MoveToPlot entity)
        )
        (setq i (1+ i))
      )
    )
  )

  (princ "\nLayer processing complete; bevel annotations and arrows moved to PLOT.")
  (princ)
)
