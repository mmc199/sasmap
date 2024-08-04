# sasmap图源管理小程序ty5-2
# 软件简介  

这个程序的主要用途在于用Excel数据表整理SASPlanet地图下载器的web图层，将zmp图层组织成符合SASPlanet软件文档中要求的文件结构。在图层数量较多时，在Excel数据表中整理、添加、修改web图层条目可以大大节约时间，管理快捷键也是在表格中更快捷。
 
软件截图

![QQ20240731-081047](https://github.com/user-attachments/assets/6e0a3805-62b6-4d5f-8ff7-706021262f2b)

由于SASPlanet软件本体只有windows的.exe发行版，linux平台都是wine运行，本插件也只提供win版，数据转换逻辑详见仓库中的ty5-2.py代码

常用图源和文档请参照：

常用奥维自定义地图（合规版）  

https://docs.qq.com/doc/DQXJLbHR1UWZYQ2hw  和  

https://www.sasgis.org/wikisasiya/doku.php/zmp  

程序本体下载地址：   

https://bitbucket.org/sas_team/sas.planet.bin/downloads/  和  

https://github.com/sasgis/sas.planet.src/releases  

sasmap图源xls及解包程序
与软件SASPlanet.exe放在同一文件夹下
zmp图源会自动生成在Maps文件夹下
汉化文件zh.mo请放在lang文件夹下

放好汉化文件后即可看到中文选项

![2024-08-05_074453](https://github.com/user-attachments/assets/0b4f4722-a74c-48ab-9aa5-3f63abc45313)

![QQ20240730-213640](https://github.com/user-attachments/assets/57360c85-10bc-4009-bf0b-82db16bae0bb)

可以给图层和叠加层定义各自的快捷键，切换更快捷

![2024-07-30_213859](https://github.com/user-attachments/assets/de2a8d0d-f7dc-402b-a6c2-7060b5b4fe23)

如果是WinInet则选择使用系统(Internet Explorer)代理设置即可正常识别
如果选择了cURL则必须要在下面手动配置代理才可以

![QQ20240730-212527](https://github.com/user-attachments/assets/664454aa-a9ed-4c57-88c0-250f08b9cdc3)

如果xls保存后未关闭则会被wps占用，程序不能读取，会报错，需要关闭xls文件

![QQ20240731-000535](https://github.com/user-attachments/assets/29ddc164-b06e-464a-8be6-fd14aa712654)

大尺寸底图拼图可以按如下配置导出，World Mercator变形小，超过65536x65536就分割，坐标文件选.w（short ext.）即.jgw。
按多边形精确裁剪，选择范围还可以导出、导入KML文件，用release里提供的脚本导入CAD后方便AL到原坐标。

![QQ20240731-084437](https://github.com/user-attachments/assets/5ab681d7-f963-41db-b5e6-4392f1613df6)

![1-20240731](https://github.com/user-attachments/assets/1bcef6da-ff97-4d44-b512-8f44b61c0ec6)

