(defun c:IPFCD (/ currentDir imgFiles img point xCoord yCoord width)
  
  ; Get the directory of the current drawing
  (setq currentDir (getvar "DWGPREFIX"))
  
  ; Get all image files from the current directory
  (setq imgFiles (append 
                  (vl-directory-files currentDir "*.jpg" 1)
                  (vl-directory-files currentDir "*.jpeg" 1)
                  (vl-directory-files currentDir "*.png" 1)
                  (vl-directory-files currentDir "*.bmp" 1)
                  ))

  ; If there are image files, loop through each and insert them
  (if imgFiles 
    (progn
      ; Initial insertion point
      (setq point (getpoint "\nSpecify initial insertion point: "))
      (setq xCoord (car point))
      (setq yCoord (cadr point))
      
      (foreach img imgFiles
        ; Insert the image at the specified point
        (command "-image" "attach" (strcat img) (list xCoord yCoord) "" "")


        ; Update the x coordinate for the next image by adding the width and the additional 1000 units
        (setq xCoord (+ xCoord 1000))
      )
    )
  )
  (princ)
)

(princ "\nInsertPicsFromCurrentDir. Type IPFCD to run.")
(princ)

;; 启动函数

(c:IPFCD)

