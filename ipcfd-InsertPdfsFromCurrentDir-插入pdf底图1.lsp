(defun c:IPFCDpdf (/ currentDir pdfFiles point xCoord yCoord)
  ; Get the directory of the current drawing
  (setq currentDir (getvar "DWGPREFIX"))
  
  ; Get all PDF files from the current directory
  (setq pdfFiles (vl-directory-files currentDir "*.pdf" 1))

  ; If there are PDF files, loop through each and insert them
  (if pdfFiles 
    (progn
      ; Initial insertion point
      (setq point (getpoint "\nSpecify initial insertion point: "))
      (setq xCoord (car point))
      (setq yCoord (cadr point))
      
      (foreach pdf pdfFiles
        ; Insert the PDF reference at the specified point
        (command "-pdfattach" pdf 1 (list xCoord yCoord) 1 0)

        ; Update the x coordinate for the next PDF by adding the width and the additional 1000 units
        (setq xCoord (+ xCoord 3000))
      )
    )
  )
  (princ)
)

(princ "\nInsertPdfsFromCurrentDir. Type IPFCD to run.")
(princ)

;; 启动函数

(c:IPFCDpdf)


