#这是一个专用工具
#Version 1.1
#By aktomleader
#2018-08-04
#Python v3.6.0:41df79263a11
import openpyxl

# 加载工作簿
NA62 = openpyxl.load_workbook(r"D:\NA62.xlsx")
print("Load 100%")
sheetnames = NA62.sheetnames
print("表信息：\n",sheetnames)

# 获取sheet名称
Sheetname_Space = sheetnames[0]
Sheetname_Device = sheetnames[1]  #设备表
Sheetname_Dot = sheetnames[2]     #点位表
# 获取sheet
Sheet_Space  = NA62[Sheetname_Space]
Sheet_Device = NA62[Sheetname_Device]
Sheet_Dot    = NA62[Sheetname_Dot]


SheetDotC_offset_column    = 3    # C列：3
SheetDeviceM_offset_column = 13   # M列：13


Sheet_Dot_C_maxcolumnNum    = 5+1  #设置待处理行数

#配置点位绝对路径
temp_SheetDotC_offset_row    = 2
temp_SheetDeviceM_offset_row = 2

while temp_SheetDotC_offset_row < Sheet_Dot_C_maxcolumnNum:         #需要处理的行数：Sheet_Dot_C_maxcolumnNum 行
    temp_Dot = str(Sheet_Dot.cell(temp_SheetDotC_offset_row, SheetDotC_offset_column).value)            #取得点位所属设备标识
    temp_Device = str(Sheet_Device.cell(temp_SheetDeviceM_offset_row,SheetDeviceM_offset_column).value) #取得待对比设备表设备标识
    print(temp_SheetDotC_offset_row)          #打印当前处理的点位表行号

    if temp_Dot == temp_Device:               #点位所属设备标识与待对比设备表设备标识一致，则取得对应设备表设备绝对路径与设备标识，合成点位绝对路径并存放在对应列
        temp_new_DeviceURL1 = str(Sheet_Device.cell(temp_SheetDeviceM_offset_row,3).value)                                   #取得设备表设备绝对路径
        temp_new_DeviceURL2 = str(Sheet_Device.cell(temp_SheetDeviceM_offset_row,11).value)                                  #取得设备表设备标识
        temp_new_DeviceURL3 = '/'
        Sheet_Dot.cell(temp_SheetDotC_offset_row,6).value = temp_new_DeviceURL1 + temp_new_DeviceURL2 + temp_new_DeviceURL3  #合成点位绝对路径并存放
        temp_SheetDotC_offset_row    += 1     #进行下一行的处理
        temp_SheetDeviceM_offset_row =  2     #初始化设备表待对比行
    else:                                     #点位所属设备标识与待对比设备表设备标识不一致，则比对下一个设备标识
        temp_SheetDeviceM_offset_row += 1

#点位映射字典
Dict_10301  = {'CT变比':'1001','平均线电压_V':'1002','零序电压_V':'1003','A_电流_A':'1004','B_电流_A':'1005'}
Dict_DeviceType = {'10301':Dict_10301}

#配置点位ID
temp_SheetDotC_offset_row    = 2
temp_SheetDeviceM_offset_row = 2

while temp_SheetDotC_offset_row < Sheet_Dot_C_maxcolumnNum:         #需要处理的行数：Sheet_Dot_C_maxcolumnNum 行
    temp_Dot = str(Sheet_Dot.cell(temp_SheetDotC_offset_row, SheetDotC_offset_column).value)                 #取得点位所属设备标识
    temp_Device = str(Sheet_Device.cell(temp_SheetDeviceM_offset_row,SheetDeviceM_offset_column).value)      #取得待对比设备表设备标识
    print(temp_SheetDotC_offset_row)

    if temp_Dot == temp_Device:              #比对一致，则在字典中查找对应设备类型ID，并进一步查找点位映射的ID
        temp_DotName  = str(Sheet_Dot.cell(temp_SheetDotC_offset_row,1).value)             #取得点位表点位名称
        temp_DeviceType   = str(Sheet_Device.cell(temp_SheetDeviceM_offset_row,9).value)   #取得设备表设备类型ID
        if temp_DotName in Dict_DeviceType[temp_DeviceType]                  #此点位在字典中存在映射则查找字典取得点位ID并存放在对应列
            temp_DotID   = Dict_DeviceType[temp_DeviceType][temp_DotName]    #查找点位ID
            Sheet_Dot.cell(temp_SheetDotC_offset_row,9).value = temp_DotID   #存放点位ID
            temp_SheetDotC_offset_row    += 1                                #进行下一个点位映射
            temp_SheetDeviceM_offset_row =  2
        else                                                                 #此点位在字典中不存在映射则写：Unknown
            Sheet_Dot.cell(temp_SheetDotC_offset_row,9).value = 'Unknown'
            temp_SheetDotC_offset_row    += 1
            temp_SheetDeviceM_offset_row =  2

    else:                                  #比对不一致，则比对下一个设备标识
        temp_SheetDeviceM_offset_row += 1

NA62.save("D:\\NA62.xlsx")
