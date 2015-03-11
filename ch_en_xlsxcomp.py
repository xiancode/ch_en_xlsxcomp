#!/usr/bin/python
#encoding=utf-8

import os
import os.path
import string
from openpyxl import load_workbook
import openpyxl.cell as ce


class XlsxTableHeader:
    """
    
    """
    def __init__(self,filename):
        try:
            self.wb = load_workbook(filename)
            self.ws = self.wb.active
            #最大行
            self.max_col  = self.ws.get_highest_column()
            #最大列
            self.max_row  = self.ws.get_highest_row()
            #
            self.get_top_part_row()
            self.get_theader_row()
            self.get_btm_part_row()
            self.get_last_col()
            self.get_theader_range()
            self.col_row_range()
            self.start = True
            print "表头区域是",self.theader_range
            
        except:
            print "初始化",filename,"失败"
            self.start = False
    def get_theader_content(self):
        """
        归并表头单元格内容，获取最终归并的单元格的值
        """
        #获取表头中合并的区域
        self.get_merged_range()
        marea_list = self.get_theader_merged_ranges()
        self.merged_cells_check(self.ws,marea_list)
        #获取未合并的区域
        nmcells = self.get_not_merged_cells()
        canmcells = self.get_can_merged_range(nmcells)
        self.merge_range(self.ws,canmcells)
        self.get_content()
        
            
    def get_top_part_row(self):
        """
        获取首列单元格顶部有边框线的行号
        """
        self.top_partition_row = []
        for i in range(1,self.max_row+1):
            cur_cell = self.ws['A'+str(i)]
            if cur_cell.style.border.top.style:
                self.top_partition_row.append(i)
        return self.top_partition_row
    
    def get_theader_row(self):
        """
        根据顶部边框线来确定表头的起始、结束行号
        """
        self.theader_start_row = self.top_partition_row[0]
        self.theader_end_row   = self.top_partition_row[1]-1
    
    def get_btm_part_row(self):
        """
        获取首列单元格底部有边框线的行号
        """
        self.btm_partition_row = []
        for i in range(1,self.max_row+1):
            cur_cell = self.ws['A'+str(i)]
            if cur_cell.style.border.bottom.style:
                self.btm_partition_row.append(i)
        return self.btm_partition_row
    
    def get_last_col(self):
        """
        根据线条来确定最后一列的列号
        """
        #从表头首行、末列开始遍历，寻找单元格顶部分割线与首列一样的
        self.theader_end_col = ""
        #theader_top_line = self.ws['A'+str(self.theader_start_row)].style.border.top.style
        for i in range(1,self.max_col+1)[::-1]:
            cur_cell = self.ws[ce.get_column_letter(i)+str(self.theader_end_row+1)]
            #if cur_cell.style.border.top.style == theader_top_line:
            if cur_cell.value != None:
                self.theader_end_col = ce.get_column_letter(i)
                break
        return self.theader_end_col
    
    def get_theader_range(self):
        """
        根据表头行号和列号，确定表头区域
        """
        self.theader_range = ""
        self.theader_range = 'A'+str(self.theader_start_row)+":"+self.theader_end_col+str(self.theader_end_row)
        return self.theader_range
    
    def col_row_range(self):
        """
        获取行号和列号的范围
        """
        self.colnum_range = [self.theader_range[0],self.theader_range[3]]
        self.rownum_range = [self.theader_range[1],self.theader_range[4]]
    
    def get_merged_range(self):
        """
        获取表格中所有合并单元格的区域
        """
        self.merged_ranges = []
        all_merged_ranges = self.ws.merged_cell_ranges
        for tmp in all_merged_ranges:
            if tmp.find('XFD') == -1:
                self.merged_ranges.append(tmp)
        return self.merged_ranges
    
    def get_theader_merged_ranges(self):
        """
        获取表头中的合并区域
        """
        self.theader_merged_ranges = []
        for cell_range in self.merged_ranges:
            if ord(cell_range[0]) >= ord(self.colnum_range[0]) and ord(cell_range[3]) <= ord(self.colnum_range[1]) \
            and ord(cell_range[1]) >= ord(self.rownum_range[0]) and ord(cell_range[4]) <= ord(self.rownum_range[1]):
                self.theader_merged_ranges.append(cell_range)
        self.theader_merged_ranges.sort()
        return self.theader_merged_ranges
    
    def get_all_cells(self,cell_ranges):
        """
        根据给定的区域，获取该区域中的所有单元格
        """
        result_cells = []
        min_col = 0
        min_row = 0
        max_col = 0
        max_row = 0
        l = list(cell_ranges)
        if len(l) == 5 and l[2] == ":":
            min_col,min_row,max_col,max_row = ord(l[0]),ord(l[1]),ord(l[3]),ord(l[4])
        else:
            print "行或者列超出了限制，不能大于26列，大于9行"
        #print min_col,min_row,max_col,max_row
        for c_t in range(min_col,max_col+1):
            for r_t in range(min_row,max_row+1):
                result_cells.append(chr(c_t)+chr(r_t))
        return result_cells
    
    def get_not_merged_cells(self):
        """
        获取表头中没有合并的单元格
        """
        theader_all_cells = self.get_all_cells(self.theader_range)
        not_merged_cells_set = set(theader_all_cells)
        for cell_range in self.theader_merged_ranges:
            not_merged_cells_set = not_merged_cells_set - set(self.get_all_cells(cell_range))
        self.not_merged_cells = list(not_merged_cells_set)
        return self.not_merged_cells
    
    def get_max_range(self,cells_range):
        """
        获取包含合并单元格的最大区域
        """
        cells = []
        min_col = 0
        min_row = 0
        max_col = 0
        max_row = 0
        for cell_range in cells_range:
            if len(cell_range) == 5 and cell_range[2] == ":":
                cells += cell_range.split(":")
            else:
                print "行或者列超出了限制，不能大于26列，大于9行"
        min_col = max_col = ord(cells[0][0])
        min_row = max_row = ord(cells[0][1])
        for cell in cells:
            if ord(cell[0]) <= min_col:
                min_col = ord(cell[0])
            if ord(cell[0]) >= max_col:
                max_col = ord(cell[0])
            if ord(cell[1])<=min_row:
                min_row = ord(cell[1])
            if ord(cell[1])>=max_row:
                max_row = ord(cell[1])  
        return chr(min_col) + chr(min_row) + ":" + chr(max_col) + chr(max_row)
    
    def get_can_merged_range(self,not_merged_cells_list):
        """
        获取还可以合并的未合并单元格
        """
        result = [] 
        tmp_set = set(not_merged_cells_list)
        while tmp_set:
            cur_cell = tmp_set.pop()
            #print cur_cell
            samecol_set = set()
            #samecol_set.add(cell)
            for cell in tmp_set:
                if cur_cell[0] == cell[0]:
                    samecol_set.add(cell)
                #print cell,
            tmp_set = tmp_set - samecol_set
            samecol_set.add(cur_cell)
            samecol_list = list(samecol_set)
            #保证行的递增顺序
            samecol_list.sort()
            result.append(samecol_list)
        return result
    
    def merged_cells_check(self,ws,merged_areas_list):
        """
        合并单元格的检测,是否值在第一个单元格
        """
        for merged_area in merged_areas_list:
            merged_cells = self.get_all_cells(merged_area)
            scell_num = merged_cells[0]
            #不用把单元格值转化为str 然后在进行string.strip() ,可以直接对单元格值进行string.strip()
            if ws[scell_num].value != None and string.strip(ws[scell_num].value) !="":
                pass
            else:
                for idx in range(1,len(merged_cells)):
                    cell_num = merged_cells[idx]
                    if ws[cell_num].value != None and string.strip(ws[cell_num].value) != "":
                        ws[scell_num].value = ws[cell_num].value
                        ws[cell_num].value = None
    
    def first_cell_value(self,ws,merged_areas_list):
        """
        合并区域中的值置于第一个单元格
        """
        pass
  
    def merge_range(self,ws,can_merged_range):
        """
        
        """
        for item in can_merged_range:
            if len(item) == 1:
                continue;
            #print item[0]+":"+item[-1]
            merge_range = self.get_max_range([item[0]+":"+item[-1]])
            #判断列是否相同
            if merge_range[0] == merge_range[3]:
                all_cells = self.get_all_cells(merge_range)
                print "all_cells",all_cells
                #
                for cell in all_cells:
                    if ws[cell].value == None:
                        ws[cell].value = ""
                if len(all_cells) >= 2:
                    #只有两行，判断第一行的底部 和 第二行的顶部是否有边框，没有边框则合并
                    if len(all_cells) == 2:
                        if ws[all_cells[0]].style.border.top.style and ws[all_cells[-1]].style.border.bottom.style \
                            and ws[all_cells[0]].style.border.bottom.style == None  \
                            and ws[all_cells[-1]].style.border.top.style == None:
                            print "合并前",ws[all_cells[0]].value,ws[all_cells[-1]].value
                            ws[all_cells[0]].value = ws[all_cells[0]].value  + " " + ws[all_cells[-1]].value
                            ws[all_cells[-1]].value = None
                            print all_cells,"单元格值合并成功"
                            print "合并后",ws[all_cells[0]].value,ws[all_cells[-1]].value
                        else:
                            print all_cells,"单元格值不能合并"
                    #超过两行，判断中间行是否包含边框
                    else:
                        if ws[all_cells[0]].style.border.top.style and ws[all_cells[-1]].style.border.bottom.style:
                            merge_flag = 1
                            for idx in range(1,len(all_cells)-1):
                                #如果有一个单元格存在上下边框线，则不能合并
                                if ws[all_cells[idx]].style.border.top.style or ws[all_cells[idx]].style.border.bottom.style :
                                    merge_flag = 0;
                                    print all_cells,"单元格值不能合并，有边框"
                                    break
                            if merge_flag ==1 :
                                for idx in range(1,len(all_cells))[::-1]:
                                    ws[all_cells[idx-1]].value = ws[all_cells[idx-1]].value + " " + ws[all_cells[idx]].value 
                                    ws[all_cells[idx]].value  = ""
                                print all_cells,"单元格值合并成功！"   
                else:
                    print "行数少于两行 error"
    
    def get_content(self):
        """
        获取单元格内容，按照二维数组存储每个单元格的内容
        """
        self.content = []
        for i in range(ord(self.rownum_range[0]),ord(self.rownum_range[1])+1):
            tmp_list = []
            for j in range(ord(self.colnum_range[0]),ord(self.colnum_range[1])+1):
                tmp_list.append(self.ws[chr(j)+chr(i)].value)
            self.content.append(tmp_list)
        return self.content
        

if __name__ == "__main__":
    ch_xl = XlsxTableHeader("B0401c.xlsx")
    if ch_xl.start:
        ch_xl.get_theader_content()
    for line in ch_xl.content:
        for item in line:
            print item,"|",
        print "-->"
    print "---"
    
    
    
        
    
        
    
    