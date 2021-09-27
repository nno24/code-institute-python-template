import xlrd

#Must be specified by user
id_col = 0                  #ID column
xpos_col = 1                #X possition column
ypos_col = 2                #Y possition column
exportFile = "sample.txt"   #export file name
workbook = xlrd.open_workbook('sample.xls')    #Your workbook
#End must be specified by user

row_cnt = 0
id_prev = 0
worksheet = workbook.sheet_by_index(0)
n_rows=worksheet.nrows
n_parts = 0
print("Number of rows: ", n_rows)
print("Number of nparts: ", n_parts)

def set_nparts(nparts):
    e7_s = '  <nParts>'
    e7_nparts = nparts
    e7_e = "</nParts>"

def write_to_file(string_to_file):
    f = open(exportFile, "a")
    f.write(string_to_file)
    f.write("\n")
    f.close()

def create_map_layer(name,nparts):
    e1 = '<mapLayer type="STANDARD_MAP" group="MAP">'
    e2 = '  <uuid>{b019a875-177f-4457-b2ab-0ea3b1406ddh}</uuid>'
    e3_s = '  <name>'
    e3_name = name
    e3_e = '</name>'
    e3 = str(e3_s) + str(e3_name) + str(e3_e)
    e4 = '  <rangeName>Ground</rangeName>'
    e5 = '  <priority>100</priority>'
    e6 = '  <overlayindex>-1</overlayindex>'
    e7 = '  <nParts></nParts>'
    e8 = '  <linecolor>27</linecolor>'
    e9 = '  <fillcolor>0</fillcolor>'
    e10 = '  <fontindex>1</fontindex>'
    e11 = '  <linestyle>0</linestyle>'
    e12 = '  <linewidth>0</linewidth>'
    print(e1)
    print(e2)
    print(e3)
    print(e4)
    print(e5)
    print(e6)
    print(e7)
    print(e8)
    print(e9)
    print(e10)
    print(e11)
    print(e12)

    write_to_file(e1)
    write_to_file(e2)
    write_to_file(e3)
    write_to_file(e4)
    write_to_file(e5)
    write_to_file(e6)
    write_to_file(e7)
    write_to_file(e8)
    write_to_file(e9)
    write_to_file(e10)
    write_to_file(e11)
    write_to_file(e12)
    write_to_file("")


def convert_polyline_to_area_file(x1,y1,x2,y2):
    e1 = '<mapPart type="MapPolyline">'
    e2 = "  <isClosed>false</isClosed>"
    e3 = "  <nPoints>2</nPoints>"
    e3_s = "  <pos>"
    e3_x1 = x1
    e3_ws = " "
    e3_y1 = y1
    e3_e = "</pos>"
    e3 = str(e3_s) + str(e3_x1) + str(e3_ws)+ str(e3_y1) + str(e3_e)
    e4_s = "  <pos>"
    e4_x2 = x2
    e4_ws = " "
    e4_y2 = y2
    e4_e = "</pos>"
    e4=str(e4_s) + str(e4_x2) + str(e4_ws) + str(e4_y2) + str(e4_e)
    e5 = "</mapPart>"
    print(e1)
    print(e2)
    print(e3)
    print(e4)
    print(e5)

    write_to_file(e1)
    write_to_file(e2)
    write_to_file(e3)
    write_to_file(e4)
    write_to_file(e5)
    write_to_file("")

def convert_mapTxt_to_area_file(x1, y1, map_text):
    e1 = '<mapPart type="MapText">'
    e2_s = "  <pos>"
    e2_x1 = x1
    e2_ws = " "
    e2_y1 = y1
    e2_e = "</pos>"
    e2 = str(e2_s) + str(e2_x1) + str(e2_ws) + str(e2_y1) + str(e2_e)
    e3_s = "  <txt>"
    e3_txt = map_text
    e3_e = "</txt>"
    e3 = str(e3_s) + str(e3_txt) + str(e3_e)
    e4 = "</mapPart>"
    print(e1)
    print(e2)
    print(e3)
    print(e4)
    write_to_file(e1)
    write_to_file(e2)
    write_to_file(e3)
    write_to_file(e4)
    write_to_file("")


create_map_layer("TEST1",n_parts)

for row_cnt in range(0,n_rows):
    id_current = worksheet.cell(row_cnt, id_col).value
    x_current = worksheet.cell(row_cnt, xpos_col).value
    y_current = worksheet.cell(row_cnt, ypos_col).value
    #Get values of first row - when new ID - Only initially
    if row_cnt == 0:
        id_sb_p1 = id_current
        x_sb_p1 = x_current
        y_sb_p1 = y_current
    #Set first and second possition of an ID
    if id_current != id_prev and id_prev != 0:
        id_sb_p2 = id_prev
        x_sb_p2 =  x_prev
        y_sb_p2 = y_prev
        # Make text label on pt.1 to an area file
        convert_mapTxt_to_area_file(x_sb_p2, y_sb_p2, id_sb_p2)
        n_parts+=1
        # Make the polyline to an area file - consisting of two points
        convert_polyline_to_area_file(x_sb_p1, y_sb_p1, x_sb_p2, y_sb_p2)
        n_parts+=1

        #Set first possition of next ID
        id_sb_p1 = id_current
        x_sb_p1 = x_current
        y_sb_p1 = y_current
    #If last line in spread sheet
    elif row_cnt == (n_rows - 1):
        print("Reach EOF")
        id_sb_p2 = id_current
        x_sb_p2 = x_current
        y_sb_p2 = y_current
        # Make text label on pt.1 to an area file
        convert_mapTxt_to_area_file(x_sb_p2, y_sb_p2, id_sb_p2)
        n_parts+=1
        # Make the polyline to an area file - consisting of two points
        convert_polyline_to_area_file(x_sb_p1, y_sb_p1, x_sb_p2, y_sb_p2)
        n_parts+=1

        #Insert correct n_parts
        lines = open(exportFile).read().splitlines()
        e7_s = '  <nParts>'
        e7_nparts = n_parts
        e7_e = '</nParts>'
        e7 = str(e7_s) + str(e7_nparts) + str(e7_e)
        lines[6] = e7
        open(exportFile, 'w').write('\n'.join(lines))


    id_prev = id_current
    x_prev = x_current
    y_prev = y_current

    row_cnt+=1





