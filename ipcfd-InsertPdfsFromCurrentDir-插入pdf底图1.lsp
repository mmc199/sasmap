(defun c:IPFCDpdf (/ currentDir pdfFiles point xCoord yCoord pdfPath)
  ; 获取当前绘图的目录
  (setq currentDir (getvar "DWGPREFIX"))
  
  ; 获取当前目录下所有 PDF 文件
  (setq pdfFiles (vl-directory-files currentDir "*.pdf" 1))

  ; 如果有 PDF 文件，循环遍历每个文件并插入
  (if pdfFiles 
    (progn
      ; 初始插入点
      (setq point (getpoint "\nSpecify initial insertion point: "))
      (setq xCoord (car point))
      (setq yCoord (cadr point))

      ; 遍历所有 PDF 文件
      (foreach pdf pdfFiles
        ; 拼接 PDF 文件的绝对路径
        (setq pdfPath (strcat currentDir pdf))

        ; 插入 PDF 文件
        (command "-pdfattach" pdfPath 1 (list xCoord yCoord) 1 0)

        ; 更新 x 坐标，加入宽度和额外的 1000 单位
        (setq xCoord (+ xCoord 3000))
      )
    )
  )
  (princ)
)

(princ "\nInsertPdfsFromCurrentDir. Type IPFCDpdf to run.")
(princ)

;; 启动函数
(c:IPFCDpdf)
