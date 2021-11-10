import openpyxl as xl
import openpyxl.utils as xlu
import openpyxl.styles as xls
import pywhatkit as pwt

wb = xl.load_workbook('C:/Users/Admin/Desktop/test_contact_list.xlsx')
ws2 = wb.active
mon_row = int(input('Enter the corresponding number for month. \n1.January\n2.Febrary\n3.March\n4.April\n5.May\n6.Ju'
                    'ne\n7.July\n8.August\n9.September\n10.October\n11.November\n12.December'))
if mon_row == 1:
    mon_row = 'C'
elif mon_row == 2:
    mon_row = 'D'
elif mon_row == 3:
    mon_row = 'E'
elif mon_row == 4:
    mon_row = 'F'
elif mon_row == 5:
    mon_row = 'G'
elif mon_row == 6:
    mon_row = 'H'
elif mon_row == 7:
    mon_row = 'I'
elif mon_row == 8:
    mon_row = 'J'
elif mon_row == 9:
    mon_row = 'K'
elif mon_row == 10:
    mon_row = 'L'
elif mon_row == 11:
    mon_row = 'M'
elif mon_row == 12:
    mon_row = 'N'
print(mon_row)
b = 0
for row in range(2, 7):
    mon_cell = mon_row + str(row)
    mob_no_place = 'A' + str(row)
    mob_no = ws2[mob_no_place].value
    color_hex = ws2[mon_cell].fill.start_color.index
    a = str(color_hex)

    if a == '00000000':
        Amount = ws2[mon_cell].value
        pwt.sendwhatmsg_instantly(mob_no, f'Your newspaper bill for the month of October is Rs.{Amount}.'
                                          f' Kindly PhonePe, GPay or UPI(request details) the bill on this number.'
                                          f'                                                                                        '

                                          f'                                                                                        '
                                          f'                                                                                        '
                                          f'NOTE: Please send screen shot/ transaction id after payment of '
                                          f'the Bill.', 10, True)
        b = b + 1

print(b)

# wb.save('C:/Users/Admin/Desktop/workbook1.xlsx')
