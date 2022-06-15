import openpyxl as xl
import openpyxl.utils as xlu
import openpyxl.styles as xls
import pywhatkit as pwt

wb = xl.load_workbook('C:/Users/Admin/Downloads/PaymentsSend.xlsx')
ws2 = wb.active
mon_row = int(input('Enter the corresponding number for month. \n1.January\n2.Febrary\n3.March\n4.April\n5.May\n6.Ju'
                    'ne\n7.July\n8.August\n9.September\n10.October\n11.November\n12.December'))
if mon_row == 1:
    mon_row = 'N'
    mon = 'January'
elif mon_row == 2:
    mon_row = 'O'
    mon = 'Febrary'
elif mon_row == 3:
    mon_row = 'D'
    mon = 'March'
elif mon_row == 4:
    mon_row = 'E'
    mon = 'April'
elif mon_row == 5:
    mon_row = 'F'
    mon = 'May'
elif mon_row == 6:
    mon_row = 'G'
    mon = 'June'
elif mon_row == 7:
    mon_row = 'H'
    mon = 'July'
elif mon_row == 8:
    mon_row = 'I'
    mon = 'August'
elif mon_row == 9:
    mon_row = 'J'
    mon = 'September'
elif mon_row == 10:
    mon_row = 'K'
    mon = 'October'
elif mon_row == 11:
    mon_row = 'L'
    mon = 'November'
elif mon_row == 12:
    mon_row = 'M'
    mon = 'December'
print(mon_row)
b = 0
total = 0
for row in range(2, 165):
    mon_cell = mon_row + str(row)
    mob_no_place = 'A' + str(row)
    color_hex = ws2[mon_cell].fill.start_color.index
    a = str(color_hex)
    color_hex2 = ws2[mob_no_place].fill.start_color.index
    num_color = str(color_hex2)
    if num_color == '00000000':
        mob_no = ws2[mob_no_place].value
        if a == '00000000':
            Amount = ws2[mon_cell].value
            pwt.sendwhatmsg_instantly(mob_no, f'Your newspaper bill for the month of April and {mon} is Rs.{Amount}.'
                                              f' Kindly PhonePe, GPay or UPI(request details) the bill on this number.\n'
                                              f'                                                                                        '

                                              f'                                                                                        '
                                              f'                                                                                        '
                                              f'NOTE: Please send screen shot/ transaction id after payment of '
                                              f'the Bill.', 22, True)

            b = b + 1
        else:
            pass

    else:
        pass
    if a == 'FF00B050':
        Amount = ws2[mon_cell].value
        total = total + Amount
    total_cell = mon_row + "128"

    ws2[total_cell] = total
wb.save('C:/Users/Admin/Downloads/PaymentsSend.xlsx')
print(b)

