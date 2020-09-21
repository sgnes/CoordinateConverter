import re
from collections import namedtuple
import xlrd
import tkinter as tk
import tkinter.filedialog
import logging
import logging.config


class CoordinateConverter(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()
        logging.config.fileConfig('logging.conf')
        self.logger = logging.getLogger(self.__class__.__name__)
        

    def create_widgets(self):
        self.coor = tkinter.StringVar()
        self.coor.set("Please select the coordinate file")
        self.bom = tkinter.StringVar()
        self.bom.set("Please select the BOM file")
        self.CoordinateADLab = tk.Label(self, textvariable=self.coor)
        self.CoordinateADLab.grid(row=0, column=0)
        #self.CoordinateAD_text = tk.Text(self, width=67, height=1)
        #self.CoordinateAD_text.grid(row=0, column=1)
        self.CoordinateAD_button = tk.Button(self, text="Select the Coordinate file", bg="lightblue",command=self.select_Coordinate)
        self.CoordinateAD_button.grid(row=0, column=2)

        self.bom_lab = tk.Label(self, textvariable=self.bom)
        self.bom_lab.grid(row=1, column=0)
        self.bom_button = tk.Button(self, text="Select the BOM file", bg="lightblue",command=self.select_bom)
        self.bom_button.grid(row=1, column=2)

        self.gen_button = tk.Button(self, text="Generate essemtec Coordinate files", bg="lightblue",command=self.gen_essemtec)
        self.gen_button.grid(row=2, column=0)


    def select_Coordinate(self):
        self.coor.set(tkinter.filedialog.askopenfilename(filetypes = (("coordinate files","*.txt"),("all files","*.*"))))
        self.logger.info("Coordinate file '{0}' selected!".format(self.coor.get()))

    def select_bom(self):
        self.bom.set(tkinter.filedialog.askopenfilename(filetypes = (("Bom files","*.xls;*.xls;*.xlsx;*.xlsm"),("all files","*.*"))))  
        self.logger.info("BOM file '{0}' selected!".format(self.bom.get()))      


    def gen_essemtec(self):
        CoordinateAD = namedtuple('CoordinateAD', 'refdes footprint midx midy refx refy padx pady tb rotation comment')
        CoordinateSch = namedtuple('CoordinateSch', 'refdes x y t part footprint')
        BomExt = namedtuple('BomExt', 'pos SchId SchDesc Quatity Remark SupplierId')

        reExp = r"^([a-zA-Z]*[\d]+)\s+([\S]+)\s+([\S]+)\s+([\S]+)\s+([\S]+)\s+([\S]+)\s+([\S]+)\s+([\S]+)\s+([\S]+)\s+([\S]+)\s+([\S ]+)\s+"

        top_list = []
        bottom_list = []
        bom_dict = {}
        
        with open(self.coor.get()) as coor:
            text = coor.read()
            res = re.findall(reExp, text, re.M)
            for i in res:
                refdes, footprint, midx, midy, refx,refy,padx, pady, tb, rotation, comment  = i
                if tb == "T":
                    top_list.append(CoordinateAD(*i))
                elif tb == "B":
                    bottom_list.append(CoordinateAD(*i))
                else:
                    raise ValueError

        wb = xlrd.open_workbook(self.bom.get())
        table = wb.sheet_by_name("Sheet1")#通过名称获取
        nrows = table.nrows
        for i in range(1,nrows,1):
            data = table.row_values(i, start_colx=0, end_colx=None)
            _, SupplierId, _, SchId, SchDesc, pos, Quatity, Remark = data
            for comp in pos.split(","):
                comp = comp.strip()
                bom_dict[comp] = BomExt(comp,  str(int(SchId)) if SchId != "" else "", SchDesc, Quatity, Remark, SupplierId)
                if SchId == "":
                    self.logger.warning("Componenet '{0}' don't have a Schindler ID!".format(comp))

        with open(self.coor.get()+"_bottom.txt", "w") as bottom:
            for bot in bottom_list:
                if bot.refdes in bom_dict:
                    bottom.write(",".join([bot.refdes, bom_dict[bot.refdes].SchId, bot.midx.strip("m"), bot.midy.strip("m"), "{:03}".format(int(float(bot.rotation))), bot.footprint]))
                    bottom.write("\r")
                else:
                    self.logger.warning("Componenet '{0}' in Coordinate file but not in BOM!".format(bot.refdes))

        with open(self.coor.get()+"_top.txt", "w") as top:
            for bot in top_list:
                if bot.refdes in bom_dict:
                    top.write(",".join([bot.refdes, bom_dict[bot.refdes].SchId, bot.midx.strip("m"), bot.midy.strip("m"), "{:03}".format(int(float(bot.rotation))), bot.footprint]))
                    top.write("\r")
                else:
                    self.logger.warning("Componenet '{0}' in Coordinate file but not in BOM!".format(bot.refdes))

root = tk.Tk()
root.title("Coordinate File Converter")
app = CoordinateConverter(master=root)
app.mainloop()