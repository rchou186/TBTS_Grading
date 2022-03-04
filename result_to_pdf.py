###############################################################################
### Make pdf result file                                                    ###
### Module to install: pandas, reportlab                                    ###
###############################################################################

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import inch, cm, mm
import pandas as pd

from defines import *

VDELTAMAX_THRES = 2.4
GRAY1 = 0.7             # 0.0=black, 1.0=white

def result_to_pdf(table, battery_model, version, result_filename):

    modules = get_max_module(battery_model)

    c = canvas.Canvas(f"{result_filename}.pdf", pagesize=landscape(A4), bottomup=0)  # bottomup=0, from top-left
    width, height = A4      # this is a tuple

    # print title info
    c.setFont('Helvetica', 16)
    c.drawString(15*mm, 50*mm, table.loc[0, 'Battery'])             # test number
    c.setFont('Helvetica', 12)
    c.drawString(15*mm, 60*mm, table.loc[0, 'B_SN'])                # B_SN
    c.drawString(70*mm, 60*mm, f"{table.loc[0, 'Date']}")           # date
    c.drawString(95*mm, 60*mm, f"{table.loc[0, 'Time']}")           # time
    c.drawString(125*mm, 60*mm, table.loc[0, 'Station'])            # station
    c.drawString(150*mm, 60*mm, get_model_result(battery_model))    # battery model
    c.drawString(185*mm, 60*mm, version)                            # software version
    c.drawString(225*mm, 60*mm, str(VDELTAMAX_THRES))

    # check half of pack
    if battery_model == 3:      # 18M+16M
        half_module = 18
    elif battery_model == 6:    # 34M+6M
        half_module = 34
    elif battery_model == 7:    # 20M+20M
        half_module = 20
    else:
        half_module = modules

    module_pitch = 6.5*mm
    module_width = 5.8*mm

    # first half pack
    # draw outter thick box outter_xs, outter_ys, outter_xe, outter_ye
    outter_x = 8*mm
    outter_y = 65*mm
    outter_w = (module_pitch * half_module) + (5*mm * 2)
    outter_h = 20*mm + (10*mm * 2)
    x_second_half_offset = outter_w + 1*mm
    c.setStrokeColorRGB(0.0, 0.0, 0.0)
    c.setLineWidth(2)
    c.rect(outter_x, outter_y, outter_w, outter_h, stroke=True)

    # draw outter thick box outter_x, outter_y, outter_w, outter_h
    inner_x = outter_x + 4*mm
    inner_y = outter_y + 9*mm
    inner_w = (module_pitch * half_module) + 1.3*mm
    inner_h = 20*mm + 2*mm
    c.setLineWidth(1)
    c.rect(inner_x, inner_y, inner_w, inner_h, stroke=True)

    # draw marker bar
    c.setLineWidth(3)
    c.line(outter_x+outter_w-2.5*mm, outter_y+12*mm, outter_x+outter_w-2.5*mm, outter_y+28*mm)

    # print signs
    c.setFont('Courier', 18)
    c.drawString(outter_x+6*mm, outter_y+6*mm, sign1[battery_model])
    c.drawString(outter_x+6*mm+module_pitch*(half_module-1), outter_y+6*mm, sign2[battery_model])

    # print bin
    c.setLineWidth(1)
    c.setFillColorRGB(0.0, 0.0, 0.0)            # text color
    for i in range(half_module):
        # draw each module rectangle
        module_x = outter_x + 5*mm + (module_pitch * i)
        module_y = outter_y + 10*mm
        module_w = module_width
        module_h = 20*mm
        # print module number
        c.setFont('Helvetica', 12)
        c.setFillGray(0.0)
        c.drawString(outter_x+5.8*mm+module_pitch*i, outter_y+35*mm, f"{i+1:02d}")

        if table.loc[i, 'BIN'] == 'NG':
            c.setFillGray(0.0)
            c.rect(module_x, module_y, module_w, module_h, fill=0, stroke=True)     # draw module rectangle
            c.setFont('Helvetica', 7.5)
            c.drawString((outter_x+5.8*mm+module_pitch*i), outter_y+15*mm, 'NG')
        elif table.loc[i, 'BIN'] == 'ERR' or table.loc[i, 'BIN'] == 'SER':
            c.setFillGray(0.0)
            c.rect(module_x, module_y, module_w, module_h, fill=0, stroke=True)     # draw module rectangle
            c.setFont('Helvetica', 6.5)
            c.drawString((outter_x+5.4*mm+module_pitch*i), outter_y+26*mm, 'ERR')
        else:
            c.setFillGray(GRAY1)
            c.rect(module_x, module_y, module_w, module_h, fill=1, stroke=True)     # draw module rectangle
            c.setFillGray(0.0)
            c.setFont('Helvetica', 12)
            c.drawString((outter_x+6.3*mm+module_pitch*i), outter_y+20.5*mm, table.loc[i, 'BIN'])

        # print module SN and VDeltaMax
        x_SV_string_offset = 9*mm
        y_SV_string_offset = 90*mm
        if table.loc[i, 'VDeltaMax'] >= VDELTAMAX_THRES:
            c.setFillGray(GRAY1)
            c.rect(outter_x+(x_SV_string_offset-3.7*mm)+module_pitch*i, outter_y+(y_SV_string_offset-9*mm), 4.5*mm, 10*mm, fill=1, stroke=False)
        c.setFont('Helvetica', 12)
        c.setFillGray(0.0)
        c.rotate(270)       # rotate 270 degrees
        c.drawString(-(outter_y+y_SV_string_offset), outter_x+x_SV_string_offset+module_pitch*i,
                    f"{table.loc[i, 'VDeltaMax']:1.2f}      {table.loc[i, 'M_SN']}")
        c.rotate(90)        # rotate back to 0 degree


    # second helf of pack
    if (modules-half_module != 0):
        outter_x = 8*mm + x_second_half_offset
        outter_y = 65*mm
        outter_w = (module_pitch * (modules - half_module)) + (5*mm * 2)
        outter_h = 20*mm + (10*mm * 2)
        x_second_half_offset = outter_w + 1*mm
        c.setStrokeColorRGB(0.0, 0.0, 0.0)
        c.setLineWidth(2)
        c.rect(outter_x, outter_y, outter_w, outter_h, stroke=True)

        # draw outter thick box outter_x, outter_y, outter_w, outter_h
        inner_x = outter_x + 4*mm
        inner_y = outter_y + 9*mm
        inner_w = (module_pitch * (modules - half_module)) + 1.3*mm
        inner_h = 20*mm + 2*mm
        c.setLineWidth(1)
        c.rect(inner_x, inner_y, inner_w, inner_h, stroke=True)

        # draw marker bar
        c.setLineWidth(3)
        c.line(outter_x+outter_w-2.5*mm, outter_y+12*mm, outter_x+outter_w-2.5*mm, outter_y+28*mm)

        # print signs
        c.setFont('Courier', 18)
        c.drawString(outter_x+6*mm, outter_y+6*mm, sign1[battery_model])
        c.drawString(outter_x+6*mm+module_pitch*(modules-half_module-1), outter_y+6*mm, sign2[battery_model])

        # print bin
        c.setLineWidth(1)
        c.setFillColorRGB(0.0, 0.0, 0.0)            # text color
        for i in range(half_module, modules):
            # draw each module rectangle
            module_x = outter_x + 5*mm + (module_pitch * (i - half_module))
            module_y = outter_y + 10*mm
            module_w = module_width
            module_h = 20*mm
            # print module number
            c.setFont('Helvetica', 12)
            c.setFillGray(0.0)
            c.drawString(outter_x+5.8*mm+module_pitch*(i-half_module), outter_y+35*mm, f"{i+1:02d}")

            if table.loc[i, 'BIN'] == 'NG':
                c.setFillGray(0.0)
                c.rect(module_x, module_y, module_w, module_h, fill=0, stroke=True)     # draw module rectangle
                c.setFont('Helvetica', 7.5)
                c.drawString((outter_x+5.8*mm+module_pitch*(i-half_module)), outter_y+15*mm, 'NG')
            elif table.loc[i, 'BIN'] == 'ERR' or table.loc[i, 'BIN'] == 'SER':
                c.setFillGray(0.0)
                c.rect(module_x, module_y, module_w, module_h, fill=0, stroke=True)     # draw module rectangle
                c.setFont('Helvetica', 6.5)
                c.drawString((outter_x+5.4*mm+module_pitch*(i-half_module)), outter_y+26*mm, 'ERR')
            else:
                c.setFillGray(GRAY1)
                c.rect(module_x, module_y, module_w, module_h, fill=1, stroke=True)     # draw module rectangle
                c.setFillGray(0.0)
                c.setFont('Helvetica', 12)
                c.drawString((outter_x+6.3*mm+module_pitch*(i-half_module)), outter_y+20.5*mm, table.loc[i, 'BIN'])

            # print module SN and VDeltaMax
            x_SV_string_offset = 9*mm
            y_SV_string_offset = 90*mm
            if table.loc[i, 'VDeltaMax'] >= VDELTAMAX_THRES:
                c.setFillGray(GRAY1)
                c.rect(outter_x+(x_SV_string_offset-3.7*mm)+module_pitch*(i-half_module), outter_y+(y_SV_string_offset-9*mm), 4.5*mm, 10*mm, fill=1, stroke=False)
            c.setFont('Helvetica', 12)
            c.setFillGray(0.0)
            c.rotate(270)       # rotate 270 degrees
            c.drawString(-(outter_y+y_SV_string_offset), outter_x+x_SV_string_offset+module_pitch*(i-half_module),
                        f"{table.loc[i, 'VDeltaMax']:1.2f}      {table.loc[i, 'M_SN']}")
            c.rotate(90)        # rotate back to 0 degree

    c.showPage()
    c.save()
