import datetime
import xlsxwriter #pip install XlsxWriter


def split(btmKey, ec2Key, jbnkey, orrKey, mllKey, ddkkey, vtrKey):

    total2k = 0 
    total500 = 0 
    total200 = 0 
    total100 = 0 
    total50 = 0 
    total20 = 0 
    total10 = 0 
    total20p = 0 
    total10p = 0 
    total5p = 0
    total2p = 0 
    total1p = 0

    cash2k = 0
    cash500 = 0
    cash200 = 0
    cash100 = 0
    cash50 = 0
    cash20 = 0
    cash10 = 0
    cash20p = 0
    cash10p = 0
    cash5p = 0
    cash2p = 0
    cash1p = 0

    totalSum = 0


    x = datetime.datetime.now()
    date = x.strftime("%x") 
    print(date)

    workbook = xlsxwriter.Workbook('CashSplit.xlsx')
 
    worksheet = workbook.add_worksheet()
    wrap_format = workbook.add_format({'text_wrap': True})
 
    worksheet.set_column('A:A', 15)
    worksheet.set_row(1, 40)
    worksheet.set_column('B:B', 10)
    worksheet.set_column('C:C', 10)
    worksheet.set_column('D:D', 10)
    worksheet.set_column('E:E', 10)
    worksheet.set_column('F:F', 10)
    worksheet.set_column('G:G', 10)
    worksheet.set_column('H:H', 10)
    worksheet.set_column('I:I', 10)
    worksheet.set_column('J:J', 10)
    worksheet.set_column('K:K', 10)
    worksheet.set_column('L:L', 10)
    worksheet.set_column('M:M', 10)
    worksheet.set_column('N:N', 10)
    worksheet.set_column('O:O', 10)
    worksheet.set_column('Q:Q', 10)

    worksheet.write_string('A1', 'DATE: ' + date)
    worksheet.write('A2', 'Denomination', wrap_format)
    worksheet.write('A3', '2000')
    worksheet.write('A4', '500')
    worksheet.write('A5', '200')
    worksheet.write('A6', '100')
    worksheet.write('A7', '50')
    worksheet.write('A8', '20')
    worksheet.write('A9', '10')
    worksheet.write('A10', '20p')
    worksheet.write('A11', '10p')
    worksheet.write('A12', '5')
    worksheet.write('A13', '2')
    worksheet.write('A14', '1')
    worksheet.write('A15', 'Cash counted', wrap_format)
    worksheet.write('A16', 'Less OB', wrap_format)
    worksheet.write('A17', 'Balance', wrap_format)
    worksheet.write('A18', 'Card')
    worksheet.write('A19', 'Digital', wrap_format)
    


    worksheet.write('B2', 'BTM\nNotes/coins', wrap_format)
    #keycode input values here
    j = 0
    for i in range(1, 13):
        worksheet.write('B'+str(i+2), int(btmKey[(j):(j+2)]))
        #counting logic
        if(i == 1):
            worksheet.write('C3', 2000 * int(btmKey[(j):(j+2)]))
            a = 2000 * int(btmKey[(j):(j+2)])
            total2k += int(btmKey[(j):(j+2)])
            cash2k += a
        if(i == 2):
            worksheet.write('C4', 500 * int(btmKey[(j):(j+2)]))
            b = 500 * int(btmKey[(j):(j+2)])
            total500 += int(btmKey[(j):(j+2)])
            cash500 += b
        if(i == 3):
            worksheet.write('C5', 200 * int(btmKey[(j):(j+2)]))
            c = 200 * int(btmKey[(j):(j+2)])
            total200 += int(btmKey[(j):(j+2)])
            cash200 += c
        if(i == 4):
            worksheet.write('C6', 100 * int(btmKey[(j):(j+2)]))
            d = 100 * int(btmKey[(j):(j+2)])
            total100 += int(btmKey[(j):(j+2)])
            cash100 += d
        if(i == 5):
            worksheet.write('C7', 50 * int(btmKey[(j):(j+2)]))
            e = 50 * int(btmKey[(j):(j+2)])
            total50 += int(btmKey[(j):(j+2)])
            cash50 += e 
        if(i == 6):
            worksheet.write('C8', 20 * int(btmKey[(j):(j+2)]))
            f = 20 * int(btmKey[(j):(j+2)])
            total20 += int(btmKey[(j):(j+2)])
            cash20 += f
        if(i == 7):
            worksheet.write('C9', 10 * int(btmKey[(j):(j+2)]))
            g = 10 * int(btmKey[(j):(j+2)])
            total10 += int(btmKey[(j):(j+2)])
            cash10 += g
        if(i == 8):
            worksheet.write('C10', 20 * int(btmKey[(j):(j+2)]))
            h = 20 * int(btmKey[(j):(j+2)])
            total20p += int(btmKey[(j):(j+2)])
            cash20p += h
        if(i == 9):
            worksheet.write('C11', 10 * int(btmKey[(j):(j+2)]))
            k = 10 * int(btmKey[(j):(j+2)])
            total10p += int(btmKey[(j):(j+2)])
            cash10p += k
        if(i == 10):
            worksheet.write('C12', 5 * int(btmKey[(j):(j+2)]))
            l = 5 * int(btmKey[(j):(j+2)])
            total5p += int(btmKey[(j):(j+2)])
            cash5p += l
        if(i == 11):
            worksheet.write('C13', 2 * int(btmKey[(j):(j+2)]))
            m = 2 * int(btmKey[(j):(j+2)])
            total2p += int(btmKey[(j):(j+2)])
            cash2p += m
        if(i == 12):
            worksheet.write('C14', 1 * int(btmKey[(j):(j+2)]))
            n = 1 * int(btmKey[(j):(j+2)])
            total1p += int(btmKey[(j):(j+2)]) 
            cash1p += n

        j = j+2

    sumBtm = a+b+c+d+e+f+g+h+k+l+m+n
    worksheet.write('C15', sumBtm)

    worksheet.write('C2', 'Cash\ncounted', wrap_format)

    worksheet.write('D2', 'EC2\nNotes/coins'), wrap_format
    j = 0
    for i in range(1, 13):
        worksheet.write('D'+str(i+2), int(ec2Key[(j):(j+2)]))
        #counting logic
        if(i == 1):
            worksheet.write('E3', 2000 * int(ec2Key[(j):(j+2)]))
            a = 2000 * int(ec2Key[(j):(j+2)])
            total2k += int(ec2Key[(j):(j+2)])
        if(i == 2):
            worksheet.write('E4', 500 * int(ec2Key[(j):(j+2)]))
            b = 500 * int(ec2Key[(j):(j+2)])
            total500 += int(ec2Key[(j):(j+2)])
        if(i == 3):
            worksheet.write('E5', 200 * int(ec2Key[(j):(j+2)]))
            c = 200 * int(ec2Key[(j):(j+2)])
            total200 += int(ec2Key[(j):(j+2)])
        if(i == 4):
            worksheet.write('E6', 100 * int(ec2Key[(j):(j+2)]))
            d = 100 * int(ec2Key[(j):(j+2)])
            total100 += int(ec2Key[(j):(j+2)])
        if(i == 5):
            worksheet.write('E7', 50 * int(ec2Key[(j):(j+2)]))
            e = 50 * int(ec2Key[(j):(j+2)])
            total50 += int(ec2Key[(j):(j+2)])
        if(i == 6):
            worksheet.write('E8', 20 * int(ec2Key[(j):(j+2)]))
            f = 20 * int(ec2Key[(j):(j+2)])
            total20 += int(ec2Key[(j):(j+2)])
        if(i == 7):
            worksheet.write('E9', 10 * int(ec2Key[(j):(j+2)]))
            g = 10 * int(ec2Key[(j):(j+2)])
            total10 += int(ec2Key[(j):(j+2)])
        if(i == 8):
            worksheet.write('E10', 20 * int(ec2Key[(j):(j+2)]))
            h = 20 * int(ec2Key[(j):(j+2)])
            total20p += int(ec2Key[(j):(j+2)])
        if(i == 9):
            worksheet.write('E11', 10 * int(ec2Key[(j):(j+2)]))
            k = 10 * int(ec2Key[(j):(j+2)])
            total10p += int(ec2Key[(j):(j+2)])
        if(i == 10):
            worksheet.write('E12', 5 * int(ec2Key[(j):(j+2)]))
            l = 5 * int(ec2Key[(j):(j+2)])
            total5p += int(ec2Key[(j):(j+2)])
        if(i == 11):
            worksheet.write('E13', 2 * int(ec2Key[(j):(j+2)]))
            m = 2 * int(ec2Key[(j):(j+2)])
            total2p += int(ec2Key[(j):(j+2)])
        if(i == 12):
            worksheet.write('E14', 1 * int(ec2Key[(j):(j+2)]))
            n = 1 * int(ec2Key[(j):(j+2)])
            total1p += int(ec2Key[(j):(j+2)])
        j = j+2

    sumEc2 = a+b+c+d+e+f+g+h+k+l+m+n
    worksheet.write('E15', sumEc2)

    worksheet.write('E2', 'Cash\ncounted', wrap_format)
    worksheet.write('F2', 'JBN\nNotes/coins', wrap_format)
    j = 0

    for i in range(1, 13):
        worksheet.write('F'+str(i+2), int(jbnKey[(j):(j+2)]))
        #counting logic
        if(i == 1):
            worksheet.write('G3', 2000 * int(jbnKey[(j):(j+2)]))
            a = 2000 * int(jbnKey[(j):(j+2)])
            total2k += int(jbnKey[(j):(j+2)])
        if(i == 2):
            worksheet.write('G4', 500 * int(jbnKey[(j):(j+2)]))
            b = 500 * int(jbnKey[(j):(j+2)])
            total500 += int(jbnKey[(j):(j+2)])
        if(i == 3):
            worksheet.write('G5', 200 * int(jbnKey[(j):(j+2)]))
            c = 200 * int(jbnKey[(j):(j+2)])
            total200 += int(jbnKey[(j):(j+2)])
        if(i == 4):
            worksheet.write('G6', 100 * int(jbnKey[(j):(j+2)]))
            d = 100 * int(jbnKey[(j):(j+2)])
            total100 += int(jbnKey[(j):(j+2)])
        if(i == 5):
            worksheet.write('G7', 50 * int(jbnKey[(j):(j+2)]))
            e = 50 * int(jbnKey[(j):(j+2)])
            total50 += int(jbnKey[(j):(j+2)])
        if(i == 6):
            worksheet.write('G8', 20 * int(jbnKey[(j):(j+2)]))
            f = 20 * int(jbnKey[(j):(j+2)])
            total20 += int(jbnKey[(j):(j+2)])
        if(i == 7):
            worksheet.write('G9', 10 * int(jbnKey[(j):(j+2)]))
            g = 10 * int(jbnKey[(j):(j+2)])
            total10 += int(jbnKey[(j):(j+2)])
        if(i == 8):
            worksheet.write('G10', 20 * int(jbnKey[(j):(j+2)]))
            h = 20 * int(jbnKey[(j):(j+2)])
            total20p += int(jbnKey[(j):(j+2)])
        if(i == 9):
            worksheet.write('G11', 10 * int(jbnKey[(j):(j+2)]))
            k = 10 * int(jbnKey[(j):(j+2)])
            total10p += int(jbnKey[(j):(j+2)])
        if(i == 10):
            worksheet.write('G12', 5 * int(jbnKey[(j):(j+2)]))
            l = 5 * int(jbnKey[(j):(j+2)])
            total5p += int(jbnKey[(j):(j+2)])
        if(i == 11):
            worksheet.write('G13', 2 * int(jbnKey[(j):(j+2)]))
            m = 2 * int(jbnKey[(j):(j+2)])
            total2p += int(jbnKey[(j):(j+2)])
        if(i == 12):
            worksheet.write('G14', 1 * int(jbnKey[(j):(j+2)]))
            n = 1 * int(jbnKey[(j):(j+2)])
            total1p += int(jbnKey[(j):(j+2)])
        j = j+2

    sumJbn = a+b+c+d+e+f+g+h+k+l+m+n
    worksheet.write('G15', sumJbn)

    worksheet.write('G2', 'Cash\ncounted', wrap_format)
    worksheet.write('H2', 'ORR\nNotes/coins', wrap_format)

    j = 0
    for i in range(1, 13):
        worksheet.write('H'+str(i+2), int(orrKey[(j):(j+2)]))
        #counting logic
        if(i == 1):
            worksheet.write('I3', 2000 * int(orrKey[(j):(j+2)]))
            a = 2000 * int(orrKey[(j):(j+2)])
            total2k += int(orrKey[(j):(j+2)])
        if(i == 2):
            worksheet.write('I4', 500 * int(orrKey[(j):(j+2)]))
            b = 500 * int(orrKey[(j):(j+2)])
            total500 += int(orrKey[(j):(j+2)])
        if(i == 3):
            worksheet.write('I5', 200 * int(orrKey[(j):(j+2)]))
            c = 200 * int(orrKey[(j):(j+2)])
            total200 += int(orrKey[(j):(j+2)])
        if(i == 4):
            worksheet.write('I6', 100 * int(orrKey[(j):(j+2)]))
            d = 100 * int(orrKey[(j):(j+2)])
            total100 += int(orrKey[(j):(j+2)])
        if(i == 5):
            worksheet.write('I7', 50 * int(orrKey[(j):(j+2)]))
            e = 50 * int(orrKey[(j):(j+2)])
            total50 += int(orrKey[(j):(j+2)])
        if(i == 6):
            worksheet.write('I8', 20 * int(orrKey[(j):(j+2)]))
            f = 20 * int(orrKey[(j):(j+2)])
            total20 += int(orrKey[(j):(j+2)])
        if(i == 7):
            worksheet.write('I9', 10 * int(orrKey[(j):(j+2)]))
            g = 10 * int(orrKey[(j):(j+2)])
            total10 += int(orrKey[(j):(j+2)])
        if(i == 8):
            worksheet.write('I10', 20 * int(orrKey[(j):(j+2)]))
            h = 20 * int(orrKey[(j):(j+2)])
            total20p += int(orrKey[(j):(j+2)])
        if(i == 9):
            worksheet.write('I11', 10 * int(orrKey[(j):(j+2)]))
            k = 10 * int(orrKey[(j):(j+2)])
            total10p += int(orrKey[(j):(j+2)])
        if(i == 10):
            worksheet.write('I12', 5 * int(orrKey[(j):(j+2)]))
            l = 5 * int(orrKey[(j):(j+2)])
            total5p += int(orrKey[(j):(j+2)])
        if(i == 11):
            worksheet.write('I13', 2 * int(orrKey[(j):(j+2)]))
            m = 2 * int(orrKey[(j):(j+2)])
            total2p += int(orrKey[(j):(j+2)])
        if(i == 12):
            worksheet.write('I14', 1 * int(orrKey[(j):(j+2)]))
            n = 1 * int(orrKey[(j):(j+2)])
            total1p += int(orrKey[(j):(j+2)])
        j = j+2

    sumOrr = a+b+c+d+e+f+g+h+k+l+m+n
    worksheet.write('I15', sumOrr)

    worksheet.write('I2', 'Cash\ncounted', wrap_format)
    worksheet.write('J2', 'MLL\nNotes/coins', wrap_format)

    j = 0
    for i in range(1, 13):
        worksheet.write('J'+str(i+2), int(mllKey[(j):(j+2)]))
        #counting logic
        if(i == 1):
            worksheet.write('K3', 2000 * int(mllKey[(j):(j+2)]))
            a = 2000 * int(mllKey[(j):(j+2)])
            total2k += int(mllKey[(j):(j+2)])
        if(i == 2):
            worksheet.write('K4', 500 * int(mllKey[(j):(j+2)]))
            b = 500 * int(mllKey[(j):(j+2)])
            total500 += int(mllKey[(j):(j+2)])
        if(i == 3):
            worksheet.write('K5', 200 * int(mllKey[(j):(j+2)]))
            c = 200 * int(mllKey[(j):(j+2)])
            total200 += int(mllKey[(j):(j+2)])
        if(i == 4):
            worksheet.write('K6', 100 * int(mllKey[(j):(j+2)]))
            d = 100 * int(mllKey[(j):(j+2)])
            total100 += int(mllKey[(j):(j+2)])
        if(i == 5):
            worksheet.write('K7', 50 * int(mllKey[(j):(j+2)]))
            e = 50 * int(mllKey[(j):(j+2)])
            total50 += int(mllKey[(j):(j+2)])
        if(i == 6):
            worksheet.write('K8', 20 * int(mllKey[(j):(j+2)]))
            f = 20 * int(mllKey[(j):(j+2)])
            total20 += int(mllKey[(j):(j+2)])
        if(i == 7):
            worksheet.write('K9', 10 * int(mllKey[(j):(j+2)]))
            g = 10 * int(mllKey[(j):(j+2)])
            total10 += int(mllKey[(j):(j+2)])
        if(i == 8):
            worksheet.write('K10', 20 * int(mllKey[(j):(j+2)]))
            h = 20 * int(mllKey[(j):(j+2)])
            total20p += int(mllKey[(j):(j+2)])
        if(i == 9):
            worksheet.write('K11', 10 * int(mllKey[(j):(j+2)]))
            k = 10 * int(mllKey[(j):(j+2)])
            total10p += int(mllKey[(j):(j+2)])
        if(i == 10):
            worksheet.write('K12', 5 * int(mllKey[(j):(j+2)]))
            l = 5 * int(mllKey[(j):(j+2)])
            total5p += int(mllKey[(j):(j+2)])
        if(i == 11):
            worksheet.write('K13', 2 * int(mllKey[(j):(j+2)]))
            m = 2 * int(mllKey[(j):(j+2)])
            total2p += int(mllKey[(j):(j+2)])
        if(i == 12):
            worksheet.write('K14', 1 * int(mllKey[(j):(j+2)]))
            n = 1 * int(mllKey[(j):(j+2)])
            total1p += int(mllKey[(j):(j+2)])
        j = j+2

    sumMll = a+b+c+d+e+f+g+h+k+l+m+n
    worksheet.write('K15', sumMll)

    worksheet.write('K2', 'Cash\ncounted', wrap_format)
    worksheet.write('L2', 'DDK\nNotes/coins', wrap_format)

    j = 0
    for i in range(1, 13):
        worksheet.write('L'+str(i+2), int(ddkKey[(j):(j+2)]))
        #counting logic
        if(i == 1):
            worksheet.write('M3', 2000 * int(ddkKey[(j):(j+2)]))
            a = 2000 * int(ddkKey[(j):(j+2)])
            total2k += int(ddkKey[(j):(j+2)])
        if(i == 2):
            worksheet.write('M4', 500 * int(ddkKey[(j):(j+2)]))
            b = 500 * int(ddkKey[(j):(j+2)])
            total500 += int(ddkKey[(j):(j+2)])
        if(i == 3):
            worksheet.write('M5', 200 * int(ddkKey[(j):(j+2)]))
            c = 200 * int(ddkKey[(j):(j+2)])
            total200 += int(ddkKey[(j):(j+2)])
        if(i == 4):
            worksheet.write('M6', 100 * int(ddkKey[(j):(j+2)]))
            d = 100 * int(ddkKey[(j):(j+2)])
            total100 += int(ddkKey[(j):(j+2)])
        if(i == 5):
            worksheet.write('M7', 50 * int(ddkKey[(j):(j+2)]))
            e = 50 * int(ddkKey[(j):(j+2)])
            total50 += int(ddkKey[(j):(j+2)])
        if(i == 6):
            worksheet.write('M8', 20 * int(ddkKey[(j):(j+2)]))
            f = 20 * int(ddkKey[(j):(j+2)])
            total20 += int(ddkKey[(j):(j+2)])
        if(i == 7):
            worksheet.write('M9', 10 * int(ddkKey[(j):(j+2)]))
            g = 10 * int(ddkKey[(j):(j+2)])
            total10 += int(ddkKey[(j):(j+2)])
        if(i == 8):
            worksheet.write('M10', 20 * int(ddkKey[(j):(j+2)]))
            h = 20 * int(ddkKey[(j):(j+2)])
            total20p += int(ddkKey[(j):(j+2)])
        if(i == 9):
            worksheet.write('M11', 10 * int(ddkKey[(j):(j+2)]))
            k = 10 * int(ddkKey[(j):(j+2)])
            total10p += int(ddkKey[(j):(j+2)])
        if(i == 10):
            worksheet.write('M12', 5 * int(ddkKey[(j):(j+2)]))
            l = 5 * int(ddkKey[(j):(j+2)])
            total5p += int(ddkKey[(j):(j+2)])
        if(i == 11):
            worksheet.write('M13', 2 * int(ddkKey[(j):(j+2)]))
            m = 2 * int(ddkKey[(j):(j+2)])
            total2p += int(ddkKey[(j):(j+2)])
        if(i == 12):
            worksheet.write('M14', 1 * int(ddkKey[(j):(j+2)]))
            n = 1 * int(ddkKey[(j):(j+2)])
            total1p += int(ddkKey[(j):(j+2)])
        j = j+2

    sumDdk = a+b+c+d+e+f+g+h+k+l+m+n
    worksheet.write('M15', sumDdk)

    worksheet.write('M2', 'Cash\ncounted', wrap_format)
    worksheet.write('N2', 'VTR\nNotes/coins', wrap_format)

    j = 0
    for i in range(1, 13):
        worksheet.write('N'+str(i+2), int(vtrKey[(j):(j+2)]))
        #counting logic
        if(i == 1):
            worksheet.write('O3', 2000 * int(vtrKey[(j):(j+2)]))
            a = 2000 * int(vtrKey[(j):(j+2)])
            total2k += int(vtrKey[(j):(j+2)])
        if(i == 2):
            worksheet.write('O4', 500 * int(vtrKey[(j):(j+2)]))
            b = 500 * int(vtrKey[(j):(j+2)])
            total500 += int(vtrKey[(j):(j+2)])
        if(i == 3):
            worksheet.write('O5', 200 * int(vtrKey[(j):(j+2)]))
            c = 200 * int(vtrKey[(j):(j+2)])
            total200 += int(vtrKey[(j):(j+2)])
        if(i == 4):
            worksheet.write('O6', 100 * int(vtrKey[(j):(j+2)]))
            d = 100 * int(vtrKey[(j):(j+2)])
            total100 += int(vtrKey[(j):(j+2)])
        if(i == 5):
            worksheet.write('O7', 50 * int(vtrKey[(j):(j+2)]))
            e = 50 * int(vtrKey[(j):(j+2)])
            total50 += int(vtrKey[(j):(j+2)])
        if(i == 6):
            worksheet.write('O8', 20 * int(vtrKey[(j):(j+2)]))
            f = 20 * int(vtrKey[(j):(j+2)])
            total20 += int(vtrKey[(j):(j+2)])
        if(i == 7):
            worksheet.write('O9', 10 * int(vtrKey[(j):(j+2)]))
            g = 10 * int(vtrKey[(j):(j+2)])
            total10 += int(vtrKey[(j):(j+2)])
        if(i == 8):
            worksheet.write('O10', 20 * int(vtrKey[(j):(j+2)]))
            h = 20 * int(vtrKey[(j):(j+2)])
            total20p += int(vtrKey[(j):(j+2)])
        if(i == 9):
            worksheet.write('O11', 10 * int(vtrKey[(j):(j+2)]))
            k = 10 * int(vtrKey[(j):(j+2)])
            total10p += int(vtrKey[(j):(j+2)])
        if(i == 10):
            worksheet.write('O12', 5 * int(vtrKey[(j):(j+2)]))
            l = 5 * int(vtrKey[(j):(j+2)])
            total5p += int(vtrKey[(j):(j+2)])
        if(i == 11):
            worksheet.write('O13', 2 * int(vtrKey[(j):(j+2)]))
            m = 2 * int(vtrKey[(j):(j+2)])
            total2p += int(vtrKey[(j):(j+2)])
        if(i == 12):
            worksheet.write('O14', 1 * int(vtrKey[(j):(j+2)]))
            n = 1 * int(vtrKey[(j):(j+2)])
            total1p += int(vtrKey[(j):(j+2)])
        j = j+2

    sumVtr = a+b+c+d+e+f+g+h+k+l+m+n
    worksheet.write('O15', sumVtr)

    worksheet.write('O2', 'Cash\ncounted', wrap_format)

    worksheet.write('Q1', 'Consolidated', wrap_format)
    worksheet.write('Q2', 'Total\nNotes', wrap_format)
    worksheet.write('R2', 'Total\nCash Counted', wrap_format)

    worksheet.write('Q3', total2k)
    worksheet.write('Q4', total500)
    worksheet.write('Q5', total200)
    worksheet.write('Q6', total100)
    worksheet.write('Q7', total50)
    worksheet.write('Q8', total20)
    worksheet.write('Q9', total10)
    worksheet.write('Q10', total20p)
    worksheet.write('Q11', total10p)
    worksheet.write('Q12', total5p)
    worksheet.write('Q13', total2p)
    worksheet.write('Q14', total1p)

    worksheet.write('R3', cash2k)
    worksheet.write('R4', cash500)
    worksheet.write('R5', cash200)
    worksheet.write('R6', cash100)
    worksheet.write('R7', cash50)
    worksheet.write('R8', cash20)
    worksheet.write('R9', cash10)
    worksheet.write('R10', cash20p)
    worksheet.write('R11', cash10p)
    worksheet.write('R12', cash5p)
    worksheet.write('R13', cash2p)
    worksheet.write('R14', cash1p)

    totalSum = sumBtm + sumEc2 + sumJbn + sumOrr + sumMll + sumDdk + sumVtr
    worksheet.write('R15', totalSum)

    workbook.close()

btmKey = input("BTM cash keycode:")
ec2Key = input("EC2 cash keycode:")
jbnKey = input("JBN cash keycode:")
orrKey = input("ORR cash keycode:")
mllKey = input("MLL cash keycode:")
ddkKey = input("DDK cash keycode:")
vtrKey = input("VTR cash keycode:")
split(btmKey, ec2Key, jbnKey, orrKey, mllKey, ddkKey, vtrKey)
#(btmKey, ec2Key, jbnkey, orrKey, mllKey, ddkkey, vtrKey)
#000403030106020000010000