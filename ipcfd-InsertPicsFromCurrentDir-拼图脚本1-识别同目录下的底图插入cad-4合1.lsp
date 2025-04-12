(defun c:IPFCD (/ currentDir imgFiles img point xCoord yCoord width uniqueFiles)
  
  ; 获取当前绘图的目录
  (setq currentDir (getvar "DWGPREFIX"))
  
  ; 从当前目录获取所有图像文件
  (setq imgFiles (append 
                  (vl-directory-files currentDir "*.jpg" 1)
                  (vl-directory-files currentDir "*.jpeg" 1)
                  (vl-directory-files currentDir "*.png" 1)
                  (vl-directory-files currentDir "*.bmp" 1)
                  (vl-directory-files currentDir "*.tif" 1)
                  ))

  ; 去重，使用一个临时列表
  (setq uniqueFiles '())
  (foreach img imgFiles
    (if (not (member img uniqueFiles))
      (setq uniqueFiles (cons img uniqueFiles))
    )
  )

  ; 如果有图像文件，循环遍历每个文件并插入
  (if uniqueFiles 
    (progn
      ; 初始插入点
      (setq point (getpoint "\nSpecify initial insertion point: "))
      (setq xCoord (car point))
      (setq yCoord (cadr point))
      
      (foreach img uniqueFiles
        ; 在指定点插入图像
        (command "-image" "attach" (strcat img) (list xCoord yCoord) "" "")

        ; 更新 x 坐标，为下一个图像添加宽度和额外的 1000 单位
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
