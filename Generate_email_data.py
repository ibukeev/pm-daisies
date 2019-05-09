import svgwrite
#from cairosvg import svg2png
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF, renderPM

from pathlib import Path
import math
from svgwrite import cm, mm
from xlrd import open_workbook


def polarToCartesian(centerX, centerY, radius, angleInDegrees):
  angleInRadians = math.radians(angleInDegrees)
  return centerX + (radius * math.cos(angleInRadians)), centerY + (radius * math.sin(angleInRadians));


def draw_for_person(person_record, picture_size, ts):
  name = person_record[0]
  email = person_record[1]
  print(name)
  print(email)
  file_name = '../results_svg/'+email + '.svg'
  png_file = '../results_png/'+email + '.png'
  config = Path(file_name)
  if config.is_file():
    # Store configuration file values
    file_name = '../results_svg/'+email + '_'+str(ts)+ '.svg'
    png_file = '../results_png/'+email + '_'+str(ts) + '.png'
    #print("File exists!!!!")
    #print(file_name)
  dwg = svgwrite.Drawing(file_name, profile='tiny', size = (picture_size[0], picture_size[1]))
  dwg.add(dwg.text(name, insert=(picture_size[0]*0.5,80), font_family='Helvetica',font_size="30px", text_anchor='middle'))
  dwg.add(dwg.text("PM Daisy", insert=(picture_size[0]*0.5,50), font_family='Helvetica',font_size="15px", text_anchor='middle'))
  color_free = str(person_record[3]).lower()
  if color_free in matching_colors.keys():
    colors = color_pallettes.get(matching_colors.get(color_free))
  else:
    colors = color_pallettes.get('Default')
  draw_daisy(dwg,center_x,center_y, picture_radius, leaf_width, n_leafs, colors, person_record[2])
  # write svg file to disk
  
  dwg.save()
  drawing = svg2rlg(file_name)
  renderPM.drawToFile(drawing, png_file, fmt="PNG")
  
def put_title(dwg, text, center_x, center_y, radius, angle, x_shift, y_shift, anchor):
  start_point = polarToCartesian(center_x, center_y, radius, angle)
  dwg.add(dwg.text(text, insert=(start_point[0]+x_shift,start_point[1]+y_shift),text_anchor=anchor,font_family='Helvetica', font_size="14px"))
  
def draw_daisy(dwg, center_x, center_y, dwg_radius, sector_w, n_leafs, colors, person_specs):
  radius_large = dwg_radius  - 10
  radius_mid = radius_large - 50
  radius_small = radius_mid - 50
  line_width = 1
  intersector_w = (360 - n_leafs*sector_w)/n_leafs;
  for i in range(0,n_leafs):
    specs = person_specs.get(titles[i][0])
    start_angle = -90 + i*(sector_w+intersector_w)
    end_angle = -90+(i+1)*sector_w+i*intersector_w
    x_shift = 10
    if titles[i][2] == 'end':
      x_shift = -0
    elif titles[i][2] == 'start':
      x_shift = 0
    y_shift = 0
    if ((i > n_leafs/2-3) and (i < n_leafs/2+2)):
      y_shift = 15
    #['start', 'end', 'middle']
    put_title(dwg, titles[i][0], center_x, center_y, radius_large+10, (start_angle + end_angle)*0.5, x_shift, y_shift, titles[i][2])
    draw_leaf(dwg, center_x, center_y, radius_large, start_angle, end_angle, colors[2], "none", line_width)
    if specs[2] > 0:
      draw_leaf(dwg, center_x, center_y, radius_large, start_angle, end_angle, colors[0], colors[0], line_width)
    if specs[1] > 0:
      draw_leaf(dwg, center_x, center_y, radius_mid, start_angle, end_angle, colors[1], colors[1], line_width)
    if specs[0] > 0:
      draw_leaf(dwg, center_x, center_y, radius_small, start_angle, end_angle, colors[2], colors[2], line_width)

def draw_leaf(dwg, center_x, center_y, radius, start_angle, end_angle, stroke_color, fill_color,line_width):
  start = polarToCartesian(center_x, center_y, radius, start_angle)
  end = polarToCartesian(center_x, center_y, radius, end_angle)
  s = 'M ' + str(start[0]) + ' ' + str(start[1]) + ' A ' + str(radius) + ' ' + str(radius) + ' 0 0 1 ' + str(end[0]) + ' ' + str(end[1]) + ' L ' + str(center_x) + ' ' + str(center_y) + ' Z'
  dwg.add(dwg.path(s).stroke(color=stroke_color,width=line_width).fill(fill_color))


color_pallettes = dict()
color_pallettes.update({'Green':["#CEFFCE","#84CF96","#009A31"]})
color_pallettes.update({'Purple':["#E193E4","#B354B6","#843283"]})
color_pallettes.update({'Blue':["#9BE1FB","#6699FF","#3366CC"]})
color_pallettes.update({'Pink':["#FFEAEE","#FFBAD2","#E47297"]})
color_pallettes.update({'Orange':["#FFFF66","#FFCC00","#FF9900"]})
color_pallettes.update({'Random':["#F9D08B","#F97D81","#9881F5"]})
color_pallettes.update({'Grey':["#DFDFDF","#999999","#000000"]})
color_pallettes.update({'Red':["#FFC3CE","#FF0000","#B52735"]})
color_pallettes.update({'Teal':["#A9DDD9","#23B5AF","#257E78"]})
color_pallettes.update({'Yellow':["#FFFF66","#FFCC00","#FF9900"]})
color_pallettes.update({'Same as in the article':["#E193E4","#B354B6","#843283"]})
color_pallettes.update({'Emerald':["#A9DDD9","#23B5AF","#257E78"]})
color_pallettes.update({'Orange/Yellow':["#FFFF66","#FFCC00","#FF9900"]})
color_pallettes.update({'Grey/Black':["#DFDFDF","#999999","#000000"]})
color_pallettes.update({'Surprise me!':["#F9D08B","#F97D81","#9881F5"]})
color_pallettes.update({'Default':["#E193E4","#B354B6","#843283"]})




book = open_workbook('../PM Daisy (Responses).xlsx')
sheet = book.sheet_by_index(0)

matching_sheet = open_workbook('Matching.xlsx').sheet_by_index(0)

matching_colors = dict()

for row_index in range(1,matching_sheet.nrows):
  matching_colors.update({matching_sheet.cell(row_index,0).value:matching_sheet.cell(row_index,1).value})

answer_options = ['I do it myself','I have a team or team member','I use resources external to the product']
titles = [["ENG",14,'start'],["UX",15,'start'], ["DATA ANALYTICS",16,'start'], ["PEOPLE OPS",21,'start'], ["PROGRAM MGMT",17,'start'], ["MARKETING",18,'end'], ["BUSINESS",20,'end'] , ["PARTNERSHIPS",19,'end'], ["RESEARCH",2,'end'], ["PM ARTIFACTS",13,'end']];
columns = [0,1,12,13,14,15,16,17,18,19,20,21,22,24,25]
header = []
for i in columns:
  header.append(sheet.cell(0,i).value)

PM_dict = dict()

for row_index in range(1,sheet.nrows):
#for row_index in range(2,3):
#for col_index in range(sheet.ncols):  
  email = sheet.cell(row_index,1).value
  job_context = sheet.cell(row_index,12).value
  PM_artifacts = sheet.cell(row_index,13).value
  name = sheet.cell(row_index,25).value
  ts = sheet.cell(row_index,0).value
  skills_dict = dict()
  answer_matrix = []
  for j in range(0,len(titles)):
    answer_matrix.clear()
    diy = 1 if sheet.cell(row_index,titles[j][1]).value.find(answer_options[0]) > -1 else 0
    dedicated = 1 if sheet.cell(row_index,titles[j][1]).value.find(answer_options[1]) > -1 else 0
    external = 1 if sheet.cell(row_index,titles[j][1]).value.find(answer_options[2]) > -1 else 0
    skills_dict.update({titles[j][0]: [diy,dedicated,external]})

  color = sheet.cell(row_index,24).value
  PM_dict.update({ts: [name, email,skills_dict,color]})


for item in PM_dict.keys():
  dims = 330
  person_record = PM_dict.get(item)
  #print(person_record)
  center_x = dims
  center_y = dims
  
  picture_radius = dims-130
  n_leafs = 10
  leaf_width = 25
  print(person_record)
  draw_for_person(person_record,[dims*2,dims*2],item)

