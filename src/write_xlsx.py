import xlsxwriter
import string

ALPHA = {char: int(val) for val, char in enumerate(list(string.ascii_uppercase))}

# for i, spec in enumerate(apt_specs):
#     tmp_ws = wb.add_worksheet(name=str(order[i]))
#
# wb.add_worksheet(name='raw_data')


def write_xlsx(apt_specs, order, raw_data, log):
    wb = xlsxwriter.Workbook('temp.xlsx')

    # create summary page formatting
    wb = write_summary_formatting(wb, order, log)

    # create spec pages
    wb = write_spec_pages(wb, apt_specs, order, log)

    # create raw data page

    # populate summary page with data.

    wb.close()


def write_summary_formatting(wb, order, log):
    # create summary worksheet
    ws = wb.add_worksheet(name='summary')

    # set column widths
    ws.set_column('A:A', 5)
    ws.set_column('B:B', 9)
    ws.set_column('C:C', 5)
    ws.set_column('D:F', 9)
    ws.set_column('G:I', 8)
    ws.set_column('J:K', 11)
    ws.set_column('L:L', 12)
    ws.set_column('M:M', 5)

    def format_header():
        # header row heights
        ws.set_row(1, 40)
        ws.set_row(2, 30)
        ws.set_row(3, 15)
        ws.set_row(5, 20)

        # formats for merged cells
        bold_label_format_18 = wb.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': 'white',
            'font_size': 18
        })
        bold_label_format_16 = wb.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': 'white',
            'font_size': 16
        })
        bold_label_format_11 = wb.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': 'white',
            'font_size': 11
        })
        normal_label_format_11 = wb.add_format({
            'bold': 0,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': 'white',
            'font_size': 11,
        })
        yellow_val_format_12 = wb.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': 'yellow',
            'font_size': 12
        })
        tan_val_format_11 = wb.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#FFF2CC',
            'font_size': 11
        })

        # title
        ws.merge_range('B2:L2', 'AIR PERMEATION TESTING - Summary', bold_label_format_18)

        # tested spec
        ws.merge_range('B3:D5', 'TESTED SPECS', bold_label_format_16)

        # start
        ws.merge_range('E4:F5', 'START', bold_label_format_16)

        # starting sample number label and format
        ws.write('G3', 'Starting \nSample #', normal_label_format_11)
        ws.merge_range('H3:I3', '', yellow_val_format_12)

        # start date & start time
        ws.write('E3', '', yellow_val_format_12)
        ws.write('F3', '', yellow_val_format_12)

        # months, days & hours labels and values
        ws.write('G4', '# Months', bold_label_format_11)
        ws.write('H4', '# Days', bold_label_format_11)
        ws.write('I4', '# Hours', bold_label_format_11)

        ws.write('G5', '0', tan_val_format_11)
        ws.write('H5', '0', tan_val_format_11)
        ws.write('I5', '0', tan_val_format_11)

        # pressure loss
        ws.merge_range('J3:L4', 'PRESSURE LOSS', bold_label_format_16)
        ws.write('J5', 'Difference', bold_label_format_11)
        ws.write('K5', '% Reduction', bold_label_format_11)
        ws.write('L5', 'Loss/Hr', bold_label_format_11)

    def format_spec_rows():

        separator_format = wb.add_format({
            'border': 1,
            'fg_color': '#808080',
        })
        spec_format = wb.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#F2F2F2',
            'font_size': 16
        })
        measured_format = wb.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#FBE5D6',
            'font_size': 12
        })
        measured_pl_format = wb.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#F8CBAD',
            'font_size': 12
        })
        curvefit_format = wb.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#DEEBF7',
            'font_size': 12
        })
        curvefit_pl_format = wb.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#BDD7EE',
            'font_size': 12
        })

        for i, row in enumerate(range(5, ((len(order)*3)+5), 3)):
            # set row heights
            ws.set_row(row, 5)
            ws.set_row(row+1, 20)
            ws.set_row(row+2, 20)

            # get merge ranges
            sep_merge = 'B'+str(row+1)+":L"+str(row+1)
            spec_merge = 'B'+str(row+2)+':C'+str(row+3)
            m_start_merge = 'E'+str(row+2)+":F"+str(row+2)
            c_start_merge = 'E'+str(row+3)+":F"+str(row+3)
            m_end_merge = 'G'+str(row+2)+":I"+str(row+2)
            c_end_merge = 'G'+str(row+3)+":I"+str(row+3)

            # merge cells with ranges & formats
            ws.merge_range(sep_merge, '', separator_format)
            ws.merge_range(spec_merge, order[i].upper(), spec_format)
            ws.merge_range(m_start_merge, '', measured_format)
            ws.merge_range(c_start_merge, '', curvefit_format)
            ws.merge_range(m_end_merge, '', measured_format)
            ws.merge_range(c_end_merge, '', curvefit_format)

            # get cell values
            m_label_cell = 'D'+str(row+2)
            m_difference = 'J'+str(row + 2)
            m_reduction = 'K'+str(row + 2)
            m_loss = 'L'+str(row + 2)

            c_label_cell = 'D'+str(row + 3)
            c_difference = 'J' + str(row + 3)
            c_reduction = 'K' + str(row + 3)
            c_loss = 'L' + str(row + 3)

            # write the rest of the cell formatting
            ws.write(m_label_cell, 'measured', measured_format)
            ws.write(m_difference, '', measured_pl_format)
            ws.write(m_reduction, '', measured_pl_format)
            ws.write(m_loss, '', measured_pl_format)
            ws.write(c_label_cell, 'curve fit', curvefit_format)
            ws.write(c_difference, '', curvefit_pl_format)
            ws.write(c_reduction, '', curvefit_pl_format)
            ws.write(c_loss, '', curvefit_pl_format)

    format_header()
    format_spec_rows()

    return wb


def write_spec_pages(wb, specs, order, log):
    spec_name = wb.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#00B0F0',
        'font_size': 18
    })
    labels_drk_14 = wb.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#D9D9D9',
        'font_size': 14
    })
    labels_drk_11 = wb.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#D9D9D9',
        'font_size': 11
    })
    start_labels = wb.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#F2F2F2',
        'font_size': 11
    })
    eq_label = wb.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#F2F2F2',
        'font_size': 18
    })
    a_label = wb.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#F2F2F2',
        'font_size': 18,
        'color': 'red'
    })
    b_label = wb.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#F2F2F2',
        'font_size': 18,
        'color': 'blue',
    })
    c_label = wb.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#F2F2F2',
        'font_size': 18,
        'color': 'green'
    })
    tan_label = wb.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#FFF2CC',
        'font_size': 18
    })
    data_label_blue = wb.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#DEEBF7',
        'font_size': 12
    })
    data_label_tan = wb.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#FFF2CC',
        'font_size': 12
    })
    data_label_green = wb.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#E2F0D9',
        'font_size': 12
    })
    normal_12 = wb.add_format({
        'bold': 0,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': 'white',
        'font_size': 12
    })
    tan_12 = wb.add_format({
        'bold': 0,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#FFF2CC',
        'font_size': 12
    })
    normal_11 = wb.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': 'white',
        'font_size': 11
    })
    tan_11 = wb.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#FFF2CC',
        'font_size': 11
    })

    data_format = wb.add_format({
        'num_format': '###,##0.00',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': 'white',
        'font_size': 11
    })

    def format_header(sheet, s):
        # set column widths
        sheet.set_column('A:H', 8)
        sheet.set_column('I:O', 12)
        # set row heights
        for row in range(6):
            sheet.set_row(row, 25)

        # write spec name
        sheet.merge_range('A1:H1', s.name, spec_name)

        # write main header labels
        sheet.merge_range('A2:B2', 'Test Start', labels_drk_14)
        sheet.merge_range('A3:B3', 'Equation', labels_drk_14)
        sheet.merge_range('A4:B4', 'Coefficients', labels_drk_14)
        sheet.merge_range('A5:B5', 'Standard Error', labels_drk_14)
        sheet.merge_range('A6:D6', 'APT Room Ideal Temperature Const.', labels_drk_11)
        sheet.merge_range('E6:H6', '295.37', normal_12)

        # write secondary labels
        sheet.write('C2', 'Date:', start_labels)
        sheet.write('F2', 'Sample #:', start_labels)
        sheet.merge_range('D2:E2', '', normal_12)
        sheet.merge_range('G2:H2', '', normal_12)

        # write equation
        eq_a = wb.add_format({'color': 'red', 'bold': True, 'font_size': 18})
        eq_e = wb.add_format({'color': 'black', 'bold': True, 'font_size': 18})
        eq_super_b = wb.add_format({'color': 'blue', 'bold': True, 'font_size': 18, 'font_script': 1})
        eq_super_other = wb.add_format({'color': 'black', 'bold': True, 'font_size': 18, 'font_script': 1})
        eq_c = wb.add_format({'color': 'green', 'bold': True, 'font_size': 18})

        sheet.merge_range('C3:H3', '', eq_label)
        sheet.write_rich_string(
            'C3',
            eq_a, 'a',
            eq_e, 'e',
            eq_super_other, '-',
            eq_super_b, 'b',
            eq_super_other, 'x',
            eq_e, '+',
            eq_c, 'c', eq_label
        )

        sheet.merge_range('C4:C5', 'a', a_label)
        sheet.merge_range('E4:E5', 'b', b_label)
        sheet.merge_range('G4:G5', 'c', c_label)

        # write remaining placeholders
        sheet.write('D4', '0', normal_11)
        sheet.write('F4', '0', normal_11)
        sheet.write('H4', '0', normal_11)
        sheet.write('D5', '0', tan_11)
        sheet.write('F5', '0', tan_11)
        sheet.write('H5', '0', tan_11)

        # write data headers
        sheet.merge_range('I1:I2', 'Time \n(hours)', data_label_blue)
        sheet.merge_range('J1:J2', 'Curve Fit \n(PSI)', data_label_blue)
        sheet.merge_range('K1:K2', 'Pressure \n(PSI)', data_label_blue)

        sheet.merge_range('L1:L2', 'Measured \nPressure \n(PSI)', data_label_tan)
        sheet.merge_range('M1:M2', 'Avg. Temp \n(K)', data_label_tan)

        sheet.merge_range('N1:N2', 'Lower \nConfidence\n Int.', data_label_green)
        sheet.merge_range('O1:O2', 'Upper \nConfidence\n Int.', data_label_green)




        return sheet

    def write_data(sheet, s):
        cf = s.curve_fit()
        popt = cf['popt']
        pcov = cf['pcov']
        pressure = s.pressure
        avg_temp = s.avg_temp

        for i in range(len(s.time)):
            # write coefficients
            row = str(i+2)
            sheet.write('D4', popt[0], data_format)
            sheet.write('F4', popt[1], data_format)
            sheet.write('H4', popt[2], data_format)

            sheet.write('D5', pcov[0][0], data_format)
            sheet.write('F5', pcov[1][1], data_format)
            sheet.write('H5', pcov[2][2], data_format)

            # write time data in hours
            sheet.write('I'+row, s.time[i], data_format)

            # write curve fit EQ
            sheet.write_formula('J'+row, '=D4*EXP(-F4*I'+row+')+H4', data_format)

            # write normalized pressure
            sheet.write_formula('K'+row, '=L'+row+'* E6/M'+row+'', data_format)

            # write measured pressure
            sheet.write('L'+row, pressure[i], data_format)

            # write average temperature
            sheet.write('M'+row, avg_temp[i], data_format)

        return sheet

    def write_charts(sheet, s):
        chart = wb.add_chart({'type': 'scatter'})
        chart.add_series({
            'name': 'Curve Fit',
            'categories': '='+s.name+'!$I$3:$I:$'+str(len(s.psi)+3),
            'values': '='+s.name+'!$J$3:$J:$'+str(len(s.psi)+3)
        })
        chart.set_title({'name': 'Pressure vs. Time'})
        chart.set_x_axis({'name': 'Time (hours)'})
        chart.set_y_axis({'name': 'Pressure (PSI)'})
        chart.set_style(11)

        sheet.insert_chart('A7', chart, {'x_offset': 25, 'y_offset': 10})

        return sheet

    for i, spec in enumerate(specs):
        ws = format_header(wb.add_worksheet(name=spec.name), s=spec)
        ws = write_data(ws, spec)
        ws = write_charts(ws, spec)

    return wb
