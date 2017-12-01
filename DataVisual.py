import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from collections import OrderedDict
from math import log, sqrt
import numpy as np

from bokeh.plotting import figure, show, output_file
import bokeh.palettes as palet
from bokeh.io import export_svgs

# function definitions
def innerRangeCreator(dim, begin,end,spacing):
    rang=[]
    colSpacing=(end-begin+spacing)/dim
    for i in range(dim):
        rang.append(colSpacing*(i)+begin)
    return rang
def outerRangeCreator(dim, begin,end,spacing):
    rang=[]
    colSpacing=(end-begin+spacing)/dim
    for i in range(dim):
        rang.append(colSpacing*(i+1)+begin-spacing)
    return rang
def spec_dim_gen(dim,spacing,y):
    dims=[]
    for i in range(dim):
        dims.append(y-spacing*i)
    return dims

# get data from excel file
df = pd.read_excel('Benchmark.xlsx', sheet_name='Tablo(Desktop Site)')
print("Column headings:")
print(df.columns)

#set data variables
data=df.values==1
specs=df.axes[0]
brands=df.axes[1]
num_brands=len(brands)
num_specs=len(specs)
colors=palet.viridis(len(brands))

 # sort data
sum_4brands=np.sum(data,0)
sum_4specs=np.sum(data,1)

sorted_data0=np.array(np.zeros([len(sum_4specs),len(sum_4brands)]))

sum_4brands=np.sum(data,0)
sorted_index_brands=sorted(range(len(sum_4brands)), key=lambda k: sum_4brands[k])
sorted_data0=data[:,sorted_index_brands]

sorted_data=np.array(np.zeros([len(sum_4specs),len(sum_4brands)]))

sum_4specs=np.sum(sorted_data0,1)
sorted_index_specs=sorted(range(len(sum_4specs)), key=lambda k: sum_4specs[k])
inv_sorted_index_specs=sorted_index_specs[::-1]
data=sorted_data0[inv_sorted_index_specs,:]

#randomize colors
newcol=[]
for brand in range(len(brands)):
    newcol.append(colors[sorted_index_brands[brand]])
colors=newcol

#set and initialize plot variables
width = 830
height = 800
begin=50
end=300

p = figure(plot_width=width, plot_height=height, title="",
    x_axis_type=None, y_axis_type=None,
    x_range=(-350, 770), y_range=(-550, 550),
    min_border=0, outline_line_color="black",
    background_fill_color=None, border_fill_color=None,
    toolbar_sticky=False)

p.xgrid.grid_line_color = None
p.ygrid.grid_line_color = None
p.output_backend = "svg"

# calculate angles and dimensions
big_angle = 2.0 * np.pi / (len(sum_4specs))
angles=[]
for spec in range(num_specs):
    angles.append(np.pi/2 - big_angle/2 - spec*big_angle)

inner_radius = innerRangeCreator(num_brands,begin,end,0)
outer_radius = outerRangeCreator(num_brands,begin,end,0)

# draw wedges
for brand in range(0,num_brands):
    brandi=num_brands-brand-1
    for spec in range(0,num_specs) :
        if(data[spec,brandi]):
            print("done")
            p.wedge(0, 0,outer_radius[brandi],
                       -big_angle+angles[spec]-big_angle*0.4, -big_angle+angles[spec]+big_angle*0.4, color=colors[brandi])
            p.wedge(0, 0,inner_radius[brandi],
                       -big_angle+angles[spec]-big_angle*0.55, -big_angle+angles[spec]+big_angle*0.55, color="white")

for spec in range(0,num_specs) :
    p.wedge(0, 0,300,
                       angles[spec]+big_angle*0.4,angles[spec]+big_angle*0.6, color="white")

p.circle(x=0,y=0,size=45,color="white")
 # arrange labels
newbra=[]
for brand in range(len(brands)):
    newbra.append(brands[sorted_index_brands[brand]])
brands=newbra

newspec=[]
for spec in range(len(specs)):
    newspec.append(specs[sorted_index_specs[spec]])
specs=newspec
#lebel brands
brand_dim=spec_dim_gen(num_brands,18,300)

p.text(360, 330, text=["Brands"],
   text_font_size="15pt", text_align="left", text_baseline="middle",text_font_style="bold")
#draw rect
for i in range(num_brands):
    p.rect(350,brand_dim[i] , width=30, height=13,
           color=colors[num_brands-i-1])
#write Brands
p.text(370, brand_dim, text=list(brands[::-1]),
   text_font_size="9pt", text_align="left", text_baseline="middle",text_font_style="bold")

#label specs
spec_dim=spec_dim_gen(num_specs,18,300)

pos=520
p.text(pos, spec_dim, text= np.arange(1,num_specs+1),
   text_font_size="9pt", text_align="left", text_baseline="middle",text_font_style="bold")
p.text(pos+20, spec_dim, text= specs[::-1],
   text_font_size="9pt", text_align="left", text_baseline="middle")
p.text(pos+30, 330, text=["Specs"] ,
   text_font_size="15pt", text_align="left", text_baseline="middle",text_font_style="bold")


# nums labels
xr = (end+20)*np.cos(np.array( angles)-big_angle)
yr = (end+20)*np.sin(np.array( angles)-big_angle)


#numbers around
p.text(xr, yr, np.arange(1,num_specs+1),angle=0,
       text_font_size="9pt", text_align="center", text_baseline="middle",text_font_style="bold")




# output part
output_file("Benchmark.html", title="Benchmark.py ")



show(p)