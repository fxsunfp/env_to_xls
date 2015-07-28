#############################################
# For:                                      #
#   环境定义     梳理整理                      #
#                                           #
# V1.3      20150723                        #
#    oop class programming                  #
#                                           #
#                                  By FX    #
#                                           #
#           版权所有        侵权不究           #
#############################################

import xlwt
import glob
import time
import os

def set_style(blod=False, color=1, bright=1):
    style = xlwt.XFStyle()

    ###颜色属性###
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = color
    
    
    ###字体属性###
    font = xlwt.Font()
    font.bold = blod
    
    ###边框属性###
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = bright
    borders.top = 1
    borders.bottom = 1

    ###文字位置属性###
    alignment = xlwt.Alignment()
    alignment.wrap = True
    alignment.vert = xlwt.Alignment.VERT_CENTER
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    
    style.pattern = pattern
    style.font = font
    style.borders = borders
    style.alignment = alignment
    
    return style

class mk_xls(object):
    def __init__(self, sheetname, ws, no=2, sep='\t'):
        self.sheetname = sheetname
        self.title = sheetname
        self.no = no
        self.row_no = no
        self.row_no_init = []
        self.col_width = []
        self.cluster = []
        self.ws = ws
        self.sep = sep
        self.start_no = 0
        self.mk_title()
        

    def mk_title(self):
        title_f = self.title + '.info'
        if os.path.exists(title_f):
            with open(title_f) as titlefile:
                for a_line in titlefile:
                    a_line = a_line.strip()
                    col_list = a_line.split()
                    self.ws.write_merge(int(col_list[1]), int(col_list[2]), int(col_list[3]), int(col_list[4]), col_list[0], set_style(blod=True, color=22))
                titlefile.close()
        else:
            self.row_no = 0
            print(self.sheetname, "title not fount")
        
    def mk_no_and_first_line(self):
        self.count = 0
        f = open(self.log, 'r')
        for a_line in f:
            if self.count == 0:
                self.first_line = a_line.rstrip('\n').split(self.sep)
            self.count += 1
        f.close
        
    def mk_start_no_and_cl_check(self):
        return True
    
    def mk_hostinfo_init(self):
        self.hostinfo_init = [self.hostinfo[1]] + self.first_line[self.start_no:]
        for x in range(len(self.hostinfo_init)):
            if len(self.row_no_init) <= x:
                self.row_no_init.append(self.row_no)
                self.col_width.append(7)
            else:
                self.row_no_init[x] = self.row_no
        
    def mk_hostinfo_auto(self, a_line):
        self.hostinfo_auto = [self.hostinfo[1]] + a_line[self.start_no:]

    def mk_row_no_and_width_init(self):
        if len(self.row_no_init) <= self.col_no:
            self.row_no_init.append(self.row_no)
            self.col_width.append(7)
        
    def mk_pre(self):
        if self.mk_start_no_and_cl_check():
            self.mk_hostinfo_init()
            return True
        else:
            return False

    def mk_bright(self):
        return 1

    def mk_col_width_extra(self):
        pass
        
    def mk_aline(self, a_line):
        self.col_no = 0
        self.a_line = a_line.rstrip('\n')
        cols = self.a_line.split(self.sep)
        self.mk_hostinfo_auto(cols)
        for a_col in self.hostinfo_auto:
            self.ws.write(self.row_no, self.col_no, a_col, set_style(bright=self.mk_bright()))
            self.mk_row_no_and_width_init()
            if len(a_col) > self.col_width[self.col_no]:
                self.col_width[self.col_no] = len(a_col)
                self.ws.col(self.col_no).width = self.col_width[self.col_no] * 320
            self.mk_col_width_extra()
            
            self.col_no += 1
    def merge_xls_extra(self):
        pass
    
    def merge_xls(self):
        if self.hostinfo_auto[1] != self.hostinfo_init[1]:
            for x in range(1,3):
                self.ws.write_merge(self.row_no_init[x], self.row_no - 1, x, x, self.hostinfo_init[x], set_style())
                self.hostinfo_init[x] = self.hostinfo_auto[x]
                self.row_no_init[x] = self.row_no
            
        if self.count == 0:
            for x in range(3):
                self.ws.write_merge(self.row_no_init[x], self.row_no, x, x, self.hostinfo_auto[x], set_style())
   
        self.merge_xls_extra()
        

    def next_p(self):
        ws.write_merge(self.row_no, self.row_no, 0, self.col_no - 1, '', set_style(color=22))
        self.row_no += 1
        
    def mk_database_title(self):
        pass
    
    def __call__(self, file, hostinfo):
        self.mk_database_title()
        self.log = file
        self.hostinfo = hostinfo
        self.title_no = 1
        self.mk_no_and_first_line()
        if self.mk_pre():
            f = open(self.log, 'r')
            for a_line in f:
                self.mk_aline(a_line)
                self.count -= 1
                self.merge_xls()
                self.row_no += 1
            f.close()
            self.next_p()
        

class mk_user_xls(mk_xls):
    def mk_bright(self):
        if self.col_no == 4:
            return 2
        return 1
        
    def mk_start_no_and_cl_check(self):
        self.ostype = self.hostinfo[3]
        if self.ostype.strip().lower() == 'aix':
            self.start_no = 2
            if self.first_line[0] != 'NONE':
                if self.first_line[0] in self.cluster:
                    if self.hostinfo[1] != self.hostinfo_init[0]:
                        self.ws.write_merge(self.row_no_init[0], self.row_no - 2, 0, 0, self.hostinfo_init[0] + ' ' + self.hostinfo[1], set_style())
                    self.ws.write_merge(self.row_no_init[0], self.row_no - 2, 1, 1, self.hostinfo_init[1] + ' ' + self.first_line[self.start_no], set_style())
                    return False
                else:
                    self.cluster.append(self.first_line[0])
        else:
            self.start_no = 0
        return True

    def mk_col_width_extra(self):
        self.ws.col(0).width = len(self.hostinfo_init[0]) * 500
        
    def merge_xls_extra(self):
        pass
        

class mk_filesystem_xls(mk_xls):
    def mk_start_no_and_cl_check(self):
        return True
    
    def mk_col_width_extra(self):
        self.ws.col(0).width = len(self.hostinfo_init[0]) * 500
    def merge_xls_extra(self):
        if self.hostinfo_auto[3] != self.hostinfo_init[3]:
            self.ws.write_merge(self.row_no_init[3], self.row_no - 1, 3, 3, self.hostinfo_init[3], set_style())
            self.hostinfo_init[3] = self.hostinfo_auto[3]
            self.row_no_init[3] = self.row_no
            
        if self.hostinfo_auto[4] != self.hostinfo_init[4]:
            for x in range(4,6):
                self.ws.write_merge(self.row_no_init[x], self.row_no - 1, x, x, self.hostinfo_init[x], set_style())
                self.hostinfo_init[x] = self.hostinfo_auto[x]
                self.row_no_init[x] = self.row_no
            
        if self.count == 0:
            for x in range(3,6):
                self.ws.write_merge(self.row_no_init[x], self.row_no, x, x, self.hostinfo_auto[x], set_style())

class mk_system_xls(mk_xls):
    def mk_start_no_and_cl_check(self):
        self.ostype = self.hostinfo[3]
        if self.ostype.strip().lower() == 'aix':
            self.start_no = 2
            if self.first_line[0] != 'NONE':
                if self.first_line[0] in self.cluster:
                    if self.hostinfo[1] != self.hostinfo_init[0]:
                        if self.hostinfo[2] == self.hostinfo_auto[1]:
                            systeminfo = self.hostinfo_init[0] + ' ' + self.hostinfo[1]
                        else:
                            systeminfo = self.hostinfo[1] + ' ' + self.hostinfo_init[0]
                        self.ws.write_merge(self.row_no_init[0], self.row_no - 2, 0, 0, self.hostinfo_init[0] + ' ' + self.hostinfo[1], set_style())

                    return False
                else:
                    self.cluster.append(self.first_line[0])
        else:
            self.start_no = 0
        return True
    
    def mk_col_width_extra(self):
        self.ws.col(0).width = len(self.hostinfo_init[0]) * 500
    def merge_xls(self):
        if self.hostinfo_auto[1] != self.hostinfo_init[1]:
            for x in range(1,9):
                self.ws.write_merge(self.row_no_init[x], self.row_no - 1, x, x, self.hostinfo_init[x], set_style())
                self.hostinfo_init[x] = self.hostinfo_auto[x]
                self.row_no_init[x] = self.row_no
            
        if self.count == 0:
            for x in range(9):
                self.ws.write_merge(self.row_no_init[x], self.row_no, x, x, self.hostinfo_auto[x], set_style())
        

    
class mk_hanode_xls(mk_xls):
    def mk_start_no_and_cl_check(self):
        self.ostype = self.hostinfo[3]
        if self.ostype.strip().lower() == 'aix':
            self.start_no = 2
            if self.first_line[0] != 'NONE':
                if self.first_line[0] in self.cluster:
                    if self.hostinfo[1] != self.hostinfo_init[0]:
                        if self.hostinfo[2] == self.hostinfo_auto[1]:
                            systeminfo = self.hostinfo_init[0] + ' ' + self.hostinfo[1]
                        else:
                            systeminfo = self.hostinfo[1] + ' ' + self.hostinfo_init[0]
                        self.ws.write_merge(self.row_no_init[0], self.row_no - 2, 0, 0, self.hostinfo_init[0] + ' ' + self.hostinfo[1], set_style())

                    return False
                else:
                    self.cluster.append(self.first_line[0])
        else:
            self.start_no = 0
        return True
    
    def mk_col_width_extra(self):
        self.ws.col(0).width = len(self.hostinfo_init[0]) * 500
        if self.hostinfo_auto[4].strip().lower() == 'aa':
            self.col_width[10] = int(len(self.hostinfo_auto[10]) / 2)
            self.col_width[11] = int(len(self.hostinfo_auto[11]) / 2)
            
        self.ws.col(10).width = self.col_width[10] * 240
        self.ws.col(11).width = self.col_width[11] * 240
        self.ws.col(3).width = 5 * 560


class mk_database_xls(mk_xls):

    def mk_database_title(self, no=0):
        if not no:
            if self.row_no:
                self.row_no += 10
        title_f = open(self.sheetname + str(no) + '.info')
        for a_line in title_f:
            self.col_no = 0
            a_line = a_line.strip()
            cols = a_line.split(self.sep)
            if len(cols) == 1:
                title_name = cols[0]
            for a_col in cols:
                self.ws.write(self.row_no, self.col_no, a_col, set_style(blod=True, color=22))
                self.col_no += 1
            self.row_no += 1
        self.ws.write_merge(self.row_no - 2, self.row_no - 2, 0, self.col_no - 1, title_name, set_style(blod=True, color=22))
        title_f.close()
            
    def mk_hostinfo_init(self):
        self.hostinfo_init = self.first_line[self.start_no:]
        for x in range(len(self.hostinfo_init)):
            if len(self.row_no_init) <= x:
                self.row_no_init.append(self.row_no)
                self.col_width.append(7)
            else:
                self.row_no_init[x] = self.row_no
    def mk_hostinfo_auto(self, a_line):
        self.hostinfo_auto = a_line[self.start_no:]

    def mk_aline(self, a_line):
        self.a_line = a_line.rstrip('\n')
        cols = self.a_line.split(self.sep)

        if cols[0] == '':
            if self.cols_width:
                self.cols_width = 0
                self.next_p()
                self.row_no += 1
                self.mk_database_title(no=self.title_no)
                self.title_no += 1
                    
            self.row_no -= 1
            return
        else:
            self.cols_width = 1
            self.col_no = 0
        self.mk_hostinfo_auto(cols)
        for a_col in self.hostinfo_auto:
            self.ws.write(self.row_no, self.col_no, a_col, set_style(bright=self.mk_bright()))
            self.mk_row_no_and_width_init()
            if len(a_col) > self.col_width[self.col_no]:
                self.col_width[self.col_no] = len(a_col)
                self.ws.col(self.col_no).width = self.col_width[self.col_no] * 320
            self.mk_col_width_extra()
            
            self.col_no += 1

    
    def mk_col_width_extra(self):
        pass
    def merge_xls(self):
        pass



if __name__ == '__main__':
    #定义每个sheet起始行数
    row_no = { '空间规划模块':2, '用户模块':2, '系统环境模块':2, '集群信息':2, '数据库模块':2}
    class_dict = { '空间规划模块':mk_filesystem_xls, '用户模块':mk_user_xls, '系统环境模块':mk_system_xls, '集群信息':mk_hanode_xls, '数据库模块':mk_database_xls}
    with open('system.info') as systemfile:
        for a_system in systemfile:
            a_system = a_system.strip()
            outfile = 'out/' + a_system + '_环境定义_' + time.strftime('%Y%m%d') + '.xls'
            wb = xlwt.Workbook()
            system_exist = 0
            log_exist = 1
            
            with open('sheet.info') as sheetfile:
                for a_sheet in sheetfile:
                    a_sheet = a_sheet.strip()
                    ws = wb.add_sheet(a_sheet, cell_overwrite_ok=True)
                    class_name = class_dict[a_sheet]
                    create_xls = class_name(a_sheet, ws, no=row_no[a_sheet])
                    
                    with open('hostlist') as hostfile:
                        for a_host in hostfile:
                            a_host = a_host.rstrip('\n')
                            hostinfo = a_host.split('\t')
                            if hostinfo[0] == a_system:
                                system_exist = 1
                                if a_sheet == "集群信息" and hostinfo[3].strip().lower() != 'aix':
                                    continue
                                logfile = a_sheet + '/' + hostinfo[2] + '.log'
                                if os.path.exists(logfile):
                                    if not log_exist:
                                        continue
                                else:
                                    print(logfile, 'not fount')
                                    log_exist = 0
                                    continue
                                
                                create_xls(logfile, hostinfo)
                            
                            
                        hostfile.close()
                sheetfile.close()

            if system_exist:
                if log_exist:
                    wb.save(outfile)
            else:
                print(a_system, 'not found')
        systemfile.close()
    
