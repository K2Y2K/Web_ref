一、填写PICS表：

step1: 打开PicsWrite_Check_v1.xlsm，点击“PicsAutoTool”菜单下的“WritePics”按钮，出现“填写PICS”窗口；
step2: 点击“选择待填写文件”从本地选择待填写的实验室提供的PICS表；点击“选择参考文件”从本地选择MTK提供的PICS表或要参考的PICS表；
step3: 点击“填写”按钮，等待几十秒，按钮变绿表示填写完毕；点击“取消”关掉窗口。

二、检查PICS表：

step1: 打开PicsWrite_Check_v1.xlsm，首先根据项目所支持情况填写每张工作表对应点的“IsSupported”列，"Supported Values"列的值会随着“IsSupported”列的值更新，
也可以手动填写“IsSupported”列的值 (注意工作表中标红的部分，需要手动整改)；
step2: 然后点击“PicsAutoTool”菜单下的“CheckPics”按钮，出现“检查PICS”窗口，同时也会在PicsWrite_Check_v1.xlsm同级路径下生成LogCheck.xlsx来记录查找和修改的数据；
step3: 点击“选择待检查的PICS文件”从本地选择所要检查的PICS表;点击下拉列表，选择项目所支持的GPRS；
step4：点击“Check”按钮，如果“IsSupported”列的数据为空，会有弹出框提示，点击“确定”则忽略提示继续Check,点击“取消”则暂停本次Check；等待一分钟左右，
待按钮变绿则表示检查完毕；点击“Cancle”关掉窗口。


注意：
PICS检查表V1.5.xls 表中红色部分是需要手动check的

