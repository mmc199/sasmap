;;;Reads world tiff file (.jgw) to scale and place image correctly in autocad.
;;;First insert all tiff images into drawing at whatever scale and insertion point.
;;;If the jgw exists in same directory and is named the same as the image selected,
;;;it will automatically be found and the image will be scaled and placed. If it is
;;;not in the same directory, the user can browse for the file.
;;;03.23.2011 Added support to create jgw files as well as support rotated images 
;;;Needs a file with 6 lines (ScaleX Rotation -Rotation -ScaleY TopLeftXCoord TopLeftYCoord)
;  « Last Edit: April 12, 2011, 09:43:43 am by ronjonp »  
(vl-load-com)


(defun my-stringp (obj)
  (= (type obj) 'STR)
)


;; 调试函数
(defun debug-print (message)
  (princ (strcat "\nDEBUG: " message))
  (princ)
)

;; 类型检查函数
(defun check-type (value expected-type)
  (cond
    ((and (eq expected-type 'str) (my-stringp value)) T)
    ((and (eq expected-type 'int) (integerp value)) T)
    ((and (eq expected-type 'real) (numberp value)) T)
    ((and (eq expected-type 'list) (listp value)) T)
    (T nil)
  )
)

;; 安全地获取关联列表值
(defun safe-assoc (key alist / result)
  (setq result (assoc key alist))
  (if result
    (cdr result)
    (progn
      (debug-print (strcat "Key not found in alist: " (vl-princ-to-string key)))
      nil
    )
  )
)

;; 将选择集转换为实体名称列表
(defun ss->list (ss / i lst)
  (if ss
    (repeat (setq i (sslength ss))
      (setq lst (cons (ssname ss (setq i (1- i))) lst))
    )
  )
  (if (null lst)
    (debug-print "Empty selection set or conversion failed")
  )
  lst
)

;; 获取文件存在的自定义函数
(defun file-exists (file-path)
  (setq file-path (vl-string-subst "\\" "/" file-path))  ; 确保使用正斜杠
  (if (not (tblsearch "LAYER" file-path))  ; 尝试打开文件
    (progn
      (setq file (open file-path "r"))  ; 尝试以只读方式打开
      (if file
        (progn
          (close file)  ; 如果成功打开，则关闭文件
          t)  ; 返回真
        nil))  ; 返回假
    nil))

;; 获取文件绝对路径的自定义函数
(defun vl-file-syst (file-path)
  (setq file-path (vl-string-subst "\\" "/" file-path))  ; 确保使用正斜杠
  (if (not (vl-string-search ":" file-path))  ; 判断是否为绝对路径
    (setq file-path (strcat (getvar "DWGPREFIX") file-path))) ; 在路径前加上当前图纸的路径
  (if (not (file-exists file-path))
    (setq file-path (strcat (getvar "TEMP") "/" file-path))) ; 如果文件不存在，尝试在临时目录查找
  file-path
)

;; 获取图像属性
(defun get-image-props (ename / obj entdata actual-path abs-path)
  (princ "\nEntering get-image-props function")
  (if (and (eq (type ename) 'ENAME)
           (= "IMAGE" (cdr (assoc 0 (setq entdata (entget ename)))))
           (setq obj (vlax-ename->vla-object ename)))
    (progn
      (princ "\nImage object found, getting properties")
      (setq actual-path (vla-get-ImageFile obj))  ; 使用 ImageFile 属性
      (setq abs-path (vl-file-syst actual-path))   ; 获取绝对路径
      (list
        (progn (princ "\nImageFile: ") (princ abs-path) abs-path)
        (progn (princ "\nInsertion point: ") (princ (cdr (assoc 10 entdata))) (cdr (assoc 10 entdata)))
        (progn (princ "\nImageWidth: ") (princ (vlax-get obj 'ImageWidth)) (vlax-get obj 'ImageWidth))
        (progn (princ "\nImageHeight: ") (princ (vlax-get obj 'ImageHeight)) (vlax-get obj 'ImageHeight))
        (progn (princ "\nU vector: ") (princ (cdr (assoc 11 entdata))) (cdr (assoc 11 entdata)))
        (progn (princ "\nV vector: ") (princ (cdr (assoc 12 entdata))) (cdr (assoc 12 entdata)))
      )
    )
    (progn
      (debug-print (strcat "Invalid image object: " (vl-princ-to-string ename)))
      nil
    )
  )
)





;; 处理单个图像
(defun process-image (ename opt pre / props name ipath opoint width height u-vec v-vec)
  (princ "\nEntering process-image function")
  (setq props (get-image-props ename))
  (if props
    (progn
      (setq name   (nth 0 props)
            opoint (nth 1 props)
            width  (nth 2 props)
            height (nth 3 props)
            u-vec  (nth 4 props)
            v-vec  (nth 5 props)
      )
      (princ "\nImage properties retrieved successfully")
      
      (if (and name (my-stringp name) (> (strlen name) 0))
        (progn
          (setq ipath (vl-filename-directory name))
          (princ (strcat "\nActual image path: " ipath))
          (if (and ipath (/= ipath ""))
            (progn
              (if (not (vl-file-directory-p ipath))
                (progn
                  (princ "\nWarning: Directory does not exist, attempting to use found location")
                  (setq ipath (vl-filename-directory (findfile name)))
                )
              )
              (if (= opt "WriteIt")
                (write-world-file ename name ipath width height u-vec v-vec opoint)
                (read-world-file ename name pre ipath u-vec v-vec opoint)
              )
              (command "_.draworder" ename "" "B")
            )
            (princ "\nError: Invalid image path")
          )
        )
        (princ "\nError: Invalid image name")
      )
    )
    (princ "\nError: Invalid image object")
  )
  (princ)
)

;; 读取世界文件
(defun read-world-file (ename name pre ipath u-vec v-vec opoint / world-file fh result scaleX rotationY rotationX scaleY topLeftX topLeftY)
  (setq world-file
        (cond
          ((= (vl-filename-extension name) ".jpg") (strcat ipath "\\" (vl-filename-base name) ".jgw"))
          ((= (vl-filename-extension name) ".png") (strcat ipath "\\" (vl-filename-base name) ".pgw"))
          ((or (= (vl-filename-extension name) ".tif") (= (vl-filename-extension name) ".tiff"))
           (strcat ipath "\\" (vl-filename-base name) ".tfw"))
          (T nil)  ; 如果没有匹配的类型，则返回 nil
        ))
  
  (if (not (findfile world-file))
    (setq world-file (getfiled "Select World File" "" (vl-filename-extension world-file) 16))
  )
  
  (if world-file
    (progn
      (setq fh (open world-file "r"))
      (if fh
        (progn
          (setq result
            (vl-catch-all-apply
              '(lambda ()
                 (setq scaleX (atof (read-line fh))
                       rotationY (atof (read-line fh))
                       rotationX (- (atof (read-line fh)))
                       scaleY (- (atof (read-line fh)))
                       topLeftX (atof (read-line fh))
                       topLeftY (atof (read-line fh)))
                 
                 ; 计算综合缩放因子
                 (setq scaleFactorX (abs (/ scaleX (car u-vec)))
                       scaleFactorY (abs (/ scaleY (cadr v-vec))))
                 (setq scaleFactor (min scaleFactorX scaleFactorY))
                                  
                 ; 缩放图像
                 (command "_.scale" ename "" opoint scaleFactor)
                 
                 ; 移动图像到正确位置
                 (command "_.move" ename "" opoint (list topLeftX topLeftY))
                 
                 ; 如果需要，旋转图像
                 (if (or (/= rotationY 0) (/= rotationX 0))
                   (command "_.rotate" ename "" opoint (rad2deg (atan rotationY scaleX)))
                 )
                 
                 ; 设置绘制顺序为底部
                 (command "_.draworder" ename "" "B")
                 
                 T  ; 返回成功标志
               )
            )
          )
          (close fh)  ; 确保文件被关闭
          (if (vl-catch-all-error-p result)
            (progn
              (princ "\nError processing world file: ")
              (princ (vl-catch-all-error-message result))
            )
            (princ (strcat "\nImage scaled and placed using " world-file))
          )
        )
        (princ (strcat "\nError: Unable to open world file: " world-file))
      )
    )
    (princ "\nWorld file not found")
  )
  (princ)
)


;; 写入世界文件
(defun write-world-file (ename name ipath width height u-vec v-vec opoint / world-file fh)
  (setq world-file
        (cond
          ((= (vl-filename-extension name) ".jpg") (strcat ipath "\\" (vl-filename-base name) ".jgw"))
          ((= (vl-filename-extension name) ".png") (strcat ipath "\\" (vl-filename-base name) ".pgw"))
          ((or (= (vl-filename-extension name) ".tif") (= (vl-filename-extension name) ".tiff"))
           (strcat ipath "\\" (vl-filename-base name) ".tfw"))
          (T nil)  ; 如果没有匹配的类型，则返回 nil
        ))
  
  (setq fh (open world-file "w"))
  (if fh
    (progn
      (write-line (rtos (car u-vec) 2 15) fh)        ; ScaleX
      (write-line (rtos (cadr u-vec) 2 15) fh)       ; Rotation
      (write-line (rtos (- (car v-vec)) 2 15) fh)    ; -Rotation
      (write-line (rtos (- (cadr v-vec)) 2 15) fh)   ; -ScaleY
      (write-line (rtos (car opoint) 2 15) fh)       ; TopLeftXCoord
      (write-line (rtos (cadr opoint) 2 15) fh)      ; TopLeftYCoord
      (close fh)
      (princ (strcat "\nWorld file created: " world-file))
    )
    (princ (strcat "\nError: Unable to create world file: " world-file))
  )
  (princ)
)






;; 主函数
(defun c:WorldFileManager ( / ss pre opt elist)
  (setq opt "ReadIt")
  (initget "ReadIt WriteIt")
  (setq opt (getkword "\nImage World File [ReadIt/WriteIt] <ReadIt>: "))
  (if (or (null opt) (= opt "")) (setq opt "ReadIt"))
  
  (princ "\nSelect image(s): ")
  (setq pre (getvar 'dwgprefix))
  (while (null elist)
    (setq ss (ssget '((0 . "IMAGE"))))
    (if ss
      (setq elist (ss->list ss))
      (princ "\nNo images selected. Please try again.")
    )
  )
  
  (if elist
    (foreach ename elist
      (process-image ename opt pre)
    )
    (princ "\nNo valid images in selection.")
  )
  (princ)
)

(princ "\nWorld File Manager tool loaded. Type WORLDFILE to run.")
(princ)

;; 启动函数
(c:WorldFileManager)



; (command "draworder" (entlast) "" "B") 


;  « Last Edit: Dec 16 2011 send to back added by alan »