import os
import xlsxwriter as xw

# #####################################################
# ### This progrma is designed to open Dental Pro's ###
# ### Ledger, Patient, and Procedur.DBF files and   ###
# ### Then crafted individualized CSVs for each     ###
# ### patient.                         - John       ###
# #####################################################

# ### Global Variables ###
Ledger = []
Patients = []
Rollerdex = []
Procedures = []
trial = []
LCheck = 0

if os.path.isdir('Ledgers'):
    pass
else:
    os.mkdir('Ledgers')

# ################################## Patient Splicer ######################################
print('Loading Dental Pro Ledger...')
with open('Core\\Patient2.csv', 'r') as front:
    for line in front:
        commas = []
        for char in range(len(line)):
            if line[char] == ',':
                commas.append(char)

        PRCN = line[:commas[0]]
        nums = "1234567890"
        chkr = len(PRCN)
        dblchk = 0
        for i in range(chkr):
            if nums.find(PRCN[i]) != -1:
                dblchk += 1
        if chkr == dblchk:
            PRCN = int(PRCN)
            Rollerdex.append(PRCN)
        else:
            pass

        PLst = line[commas[0] + 1:commas[1]]
        PSfx = line[commas[1] + 1:commas[2]]
        PFst = line[commas[2] + 1:]
        # PFst = PFst.replace(' ',"")
        PFst = PFst.replace('\n', '')

        if PSfx != '':
            fname = PFst + ' ' + PSfx
        else:
            fname = PFst

        PID = [PRCN, PLst, fname]
        Patients.append(PID)
    front.close()


# ################################ Row Reconstructuring ###################################
def launder(fund):
    nums = "1234567890"
    n_count = 0
    for num in nums:
        if fund.find(str(num)) != -1:
            n_count += 1
    if n_count == 0:
        fund = 0.0        
    else:
        fund = fund.replace('\n', '')
        fund = fund.replace(',', '')
        exc = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        for char in exc:
            if fund.find(char) != -1:
                fund = fund.replace(char, '')
        fund = float(fund)
    return fund

def CD_Audit(trans, LDesc, item03):
    Creditors = ['Insurance Check', 'Visa', 'Master Card', 'Personal Check', 'Discover', 'AmEx', 'American Express', 'Cash', 'At Visit', 'On Acct', 'On Account', 'Mailed', 'Discount','Charge Pmt','Payment by CC']
    Debtors = ['refund', 'void', 'canceled', 'reimburse']
    Creditors.sort()
    Debtors.sort()
    fund = launder(trans)

    # Check for Cred/Debt descriptors
    pros = 0
    if LDesc.find('Transfer') != -1:
        if LDesc.find(' To ') != -1:
            pros -= 1
        else:
            pros += 1
    else:
        for cred in Creditors:
            if LDesc.find(cred) != -1:
                pros += 1

        for debt in Debtors:
            if LDesc.find(debt) != -1:
                pros -= 1

    # If the line item's a credit, assign it's credit flag
    ldesc = LDesc.lower()
    if pros >= 1:
        ldesc= LDesc.lower()
        if ldesc.find('credit') != -1:
            icon = '²'
            item03 = 10003
            item09 = fund
            item10 = 0
        else:
            item09 = fund
            item10 = 0
            if ldesc.find('insurance') != -1:
                icon = 'o'
                item03 = 10000
            else:
                icon = '¡'
                item03 = 10001
    else:
        item09 = 0
        item10 = fund
        if LDesc.find('Transfer') != -1:
            icon = '²'
            item03 = 10004
        else:
            icon = ''

    return icon, item03, item09, item10

with open('Core/Ledger2.csv', 'r') as book:
    for line in book:

        LC = []
        for spot in range(len(line)):
            if line[spot] == ',':
                LC.append(spot)

        item01 = line[:LC[0]]             # PRCN
        item02 = line[LC[0] + 1:LC[1]]    # Date
        item03 = line[LC[1] + 1:LC[2]]    # ADA Code
        item04 = line[LC[2] + 1:LC[3]]    # Base Description
        item05 = line[LC[3] + 1:LC[4]]    # Tooth No
        item06 = line[LC[4] + 1:LC[5]]    # Quadrant No
        item07 = line[LC[5] + 1:LC[6]]    # Surface No
        item08 = line[LC[6] + 1:LC[7]]    # Check / CC No
        item09 = line[LC[7] + 1:]         # Transaction Amount

        # Lengthen out descriptions if applicable
        LDesc = item04
        if item08 == "C":
            if item05 != "" and item05 != "T":
                LDesc += '  ' + item05
            elif item06 != "" and item06 != "Q":
                LDesc += '  ' + item06
            elif item07 != "" and item07 != "S":
                LDesc += '  ' + item07
            else:
                pass
        else:
            LDesc = item04 + '  No. ' + item08

        # ### item09 / Credit / Debit Section
        icon, ADACode, Credit, Debit = CD_Audit(item09,item04,item03)

        # [Item No, Date, ADA Code, Description, Credit, Debit]
        newrow = [int(item01), int(item02), int(ADACode), LDesc, icon, float(Credit), float(Debit)]
        Ledger.append(newrow)
    book.close()


# #################################### Sort Ledger ########################################


def reSort(ledger):
    # new_list = sorted(a_list, key=lambda x: (len(x), x))
    # https://www.kite.com/python/answers/how-to-sort-by-two-keys-in-python
    sorter = sorted(ledger, key=lambda x: (x[1],x[2]))
    return sorter

# ################################### Ledger Rebuild ######################################


def Ledger_Export(ledger, PRCN):  # Build a Ledger Spreadsheet
    global Patients
    global file
    # Perhaps check out https://openpyxl.readthedocs.io/en/stable/worksheet_tables.html
    Headers = ['Line No', 'Date', 'Code', 'Description', 'þ', 'Credit', 'Debit', '', 'Total', '', 'Ins Ln', 'Ins Paid', 'Pt Ln 1', 'Amt 1', 'Pt Ln 2', 'Amt 2', 'Pt Ln 3', 'Amt 3', 'Run Total', ""]

    # Create Folder and PName
    for card in Patients:
        if PRCN == card[0]:
            Folder = card[1]
            PFst = card[2]

    path = 'Ledgers/' + Folder + '/'
    if os.path.isdir(path) is True:
        pass
    else:
        os.mkdir(path)

    file = path + PFst + '.xlsx'
    if PRCN < 10:
        rcn = '[000' + str(PRCN) + '] '
    elif PRCN < 100:
        rcn = '[00' + str(PRCN) + '] '
    elif PRCN < 1000:
        rcn = '[0' + str(PRCN) + '] '
    else:
        rcn = '[' + str(PRCN) + '] '
    
    print('Crafting Ledger file for ' + rcn + Folder + ', ' + PFst)
    
    # ### Create Workbook
    workbook = xw.Workbook(file)
    worksheet = workbook.add_worksheet(PFst)

    # ### Formatting
    icons = ['o','¡','²']
    bg_colors = ['#FFFFFF','#EBF1DE','#DAEEF3','#E4DFEC']
    font_colors = ['#FFFFFF','#EBF1DE','#DAEEF3','#E4DFEC','#000000']
    merge = workbook.add_format({'size': 20})
    title = workbook.add_format({'size': 20, 'bold': 1, 'align': 'center', 'valign': 'center'})
    bold = workbook.add_format({'bold': True})
    accounting = workbook.add_format({'num_format': '[$$-409]#,##0.00'})
    wingding = workbook.add_format({'font':'Wingdings','align':'center'})
    date = workbook.add_format({'num_format': 'mm/dd/yy','align':'center'})
    lineno = workbook.add_format({'align':'center'})

    # ### Set Column widths
    Col_Widths = [7, 10, 5, 36, 3, 11, 11, 3, 12, 3, 5.5, 11, 6, 11, 6, 11, 6, 11]
    CW = len(Col_Widths)
    for w in range(CW):
        worksheet.set_column(w,w,Col_Widths[w])

    # ### Name Header in Line 0
    pt_name = Folder + ', ' + PFst
    worksheet.merge_range('A1:R1', 'Merged Range',merge)
    worksheet.write(0,0,pt_name, title)

    # ### Create Headers in Line 1
    heading = len(Headers)
    for title in range(heading):
        worksheet.write(1,title,Headers[title],bold)
    worksheet.write(1,4,'þ',wingding)

    # ### Create Line 2
    worksheet.write(2,0,1,lineno)
    worksheet.write(2,1,'=B4',date)
    worksheet.write(2,3,'Starting Balance')
    worksheet.write(2,8,0.00,accounting)

    payments = 0
    for row in range(len(ledger)):
        line = ledger[row]      
        R = row + 3
        L_Date = line[1]
        L_Code = line[2]
        L_Desc = line[3]
        L_Wing = line[4]
        L_Cred = line[5]
        L_Debt = line[6]


        years = [40909,41275,41640,42005,42370,42736,43101,43466,43831,44197,44562]  # 1/1/2012-2022
        splits = []
        if row != len(ledger)-1:
            date1 = L_Date
            line2 = ledger[row+1]
            date2 = line2[1]
            if date1 == date2:
                b_num = 0
            else:
                for y in range(len(years)-1):
                    if date1 > years[y] and date1 < years[y+1]:
                        if date2 > years[y+1]:
                            b_num = 1
                            splits.append(row)
                            break
                        else:
                            b_num = 0
                    else:
                        b_num = 0

        if L_Wing in icons:  # Is this a funding line?
            payments += 1
            cfont = 4
            L_Code = ""
            if L_Wing == icons[0]:
                color = 1
                dfont = 1
            elif L_Wing == icons[1]:
                color = 2
                dfont = 2
            elif L_Wing == icons[2] and L_Debt == 0:
                color = 3
                cfont = 4
                dfont = 3
            elif L_Wing == icons[2] and L_Cred == 0:
                color = 3
                cfont = 3
                dfont = 4
            else:
                pass
        else:
            color = 0
            dfont = 4
            cfont = 0


        # ### Row formatting
        nmbr = workbook.add_format({'align':'center','bottom':b_num,'bg_color':bg_colors[color]})
        date = workbook.add_format({'num_format': 'mm/dd/yy','align':'center','bottom':b_num,'bg_color':bg_colors[color]})
        norm = workbook.add_format({'bottom':b_num,'bg_color':bg_colors[color]})
        wing = workbook.add_format({'font': 'Wingdings','align':'center','bottom':b_num,'bg_color':bg_colors[color]})
        cred = workbook.add_format({'num_format': '[$$-409]#,##0.00','bottom':b_num,'bg_color':bg_colors[color],'font_color':font_colors[cfont]})
        debt = workbook.add_format({'num_format': '[$$-409]#,##0.00','bottom':b_num,'bg_color':bg_colors[color],'font_color':font_colors[dfont]})


        # ### Write line numbers per row
        codeA = '=IF(D' + str(R) + '="","",Sum(Max(A$3:A' + str(R) + ')+1))'
        worksheet.write(R,0,codeA,nmbr)         # Line Number

        # ### Write out light items
        worksheet.write(R,1,L_Date,date)        # Date
        worksheet.write(R,2,L_Code,norm)        # Code
        worksheet.write(R,3,L_Desc,norm)        # Description
        worksheet.write(R,4,L_Wing,wing)        # Wingding Line
        worksheet.write(R,5,L_Cred,cred)        # Credit
        worksheet.write(R,6,L_Debt,debt)        # Debits
        worksheet.write(R,7,"",norm)            # Spacer

        # ### Write out the I to S Accouting column
        R1 = R + 1
        count = workbook.add_format({'num_format': '[$$-409]#,##0.00','bottom':b_num,})
        box_L = workbook.add_format({'align':'center','left':1,'bottom':b_num})
        box_B = workbook.add_format({'num_format': '[$$-409]#,##0.00','bottom':b_num})
        codeI = '=IF(A' + str(R1) + '="","",I' + str(R) + '+(F' + str(R1) + '-G' + str(R1) + '))'
        codeS = '=IF(C' + str(R1) + '="","",G' + str(R1) + '-Sum(L' + str(R1) + ',N' + str(R1) + ',P' + str(R1)  + ',R' + str(R1) + '))'
        worksheet.write(R,8,codeI,count)        # Add up section
        worksheet.write(R,9,"",count)            # Spacer
        worksheet.write(R,10,"",box_L)          # Insurance Payment Line
        worksheet.write(R,11,"",box_B)          # Insurance Payment Amount
        worksheet.write(R,12,"",box_L)          # Patient Payment Line     1
        worksheet.write(R,13,"",box_B)          # Patient Payment Amount   1 
        worksheet.write(R,14,"",box_L)          # Patient Payment Line     2
        worksheet.write(R,15,"",box_B)          # Patient Payment Amount   2
        worksheet.write(R,16,"",box_L)          # Patient Payment Line     3
        worksheet.write(R,17,"",box_B)          # Patient Payment Amount   3
        worksheet.write(R,18,codeS,count)       # Line Item cost remainder


    # ########################## After Ledger Going Down ##################################
    # ### Picking up where we left off on the ledger
    R += 1   # Where the ledger left off
    t_rows = len(ledger) + int((len(ledger)/50)*20)  # Where the program will keep going

    # ### A through S AL formats
    lineno = workbook.add_format({'align':'center'})
    date = workbook.add_format({'num_format': 'mm/dd/yy','align':'center'})
    accounting = workbook.add_format({'num_format': '[$$-409]#,##0.00'})
    box_L = workbook.add_format({'align':'center','left':1})
    box_B = workbook.add_format({'num_format': '[$$-409]#,##0.00'})

    # ### A through S AL cell formulas
    codeA = '=IF(D' + str(R1) + '="","",Sum(Max(A$3:A' + str(R) + ')+1))'
    codeD = '=IF(D' + str(R1) + '="","",C' + str(R) + ')'
    codeI = '=IF(A' + str(R1) + '="","",I' + str(R) + '+(F' + str(R1) + '-G' + str(R1) + '))'
    codeS = '=IF(C' + str(R1) + '="","",G' + str(R1) + '-Sum(L' + str(R1) + ',N' + str(R1) + ',P' + str(R1)  + ',R' + str(R1) + '))'

    # ### write the lines
    while R < t_rows:    
        R1 = R + 1

        # ### A through S AL cell formulas
        codeA = '=IF(D' + str(R1) + '="","",Sum(Max(A$3:A' + str(R) + ')+1))'
        codeD = '=IF(D' + str(R1) + '="","",B' + str(R) + ')'
        codeI = '=IF(A' + str(R1) + '="","",I' + str(R) + '+(F' + str(R1) + '-G' + str(R1) + '))'
        codeS = '=IF(C' + str(R1) + '="","",G' + str(R1) + '-Sum(L' + str(R1) + ',N' + str(R1) + ',P' + str(R1)  + ',R' + str(R1) + '))'

        # ### Write A-S lines
        worksheet.write(R,0,codeA,lineno)       # Line number
        worksheet.write(R,1,codeD,date)         # Date
        worksheet.write(R,8,codeI,accounting)   # Accurment
        worksheet.write(R,10,"",box_L)          # Insurance Payment Line
        worksheet.write(R,11,"",box_B)          # Insurance Payment Amount
        worksheet.write(R,12,"",box_L)          # Patient Payment Line     1
        worksheet.write(R,13,"",box_B)          # Patient Payment Amount   1 
        worksheet.write(R,14,"",box_L)          # Patient Payment Line     2
        worksheet.write(R,15,"",box_B)          # Patient Payment Amount   2
        worksheet.write(R,16,"",box_L)          # Patient Payment Line     3
        worksheet.write(R,17,"",box_B)          # Patient Payment Amount   3
        worksheet.write(R,18,codeS,count)       # Line Item cost remainder
        R += 1

    # ############################ Payment Line Finder ####################################
    spread = 20 + payments              # Number of columns before payment columns + number of payments
    spread_pad = spread + 10           # Add additional payment columns
    Edge1 = spread_pad + 1           # Checking Column 1
    Edge2 = spread_pad + 2           # Checking Column 2
    column = ['A','B','C','D','E','F','G','H','I','','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z',
        'AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ',
        'BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY','BZ',
        'CA','CB','CC','CD','CE','CF','CG','CH','CI','CJ','CK','CL','CM','CN','CO','CP','CQ','CR','CS','CT','CU','CV','CW','CX','CY','CZ',
        'DA','DB','DC','DD','DE','DF','DG','DH','DI','DJ','DK','DL','DM','DN','DO','DP','DQ','DR','DS','DT','DU','DV','DW','DX','DY','DZ',]

    col_cap = Edge1 - 26

    # ### Cell for Max Number of Payments
    #print('I have ' + str(len(column)) + ' columns, with Edge1 at ' + str(Edge1))
    MAX = '=Max(' + column[Edge2] + '3:' + column[Edge2] + str(t_rows) + ')'
    worksheet.write(0,Edge1,MAX)  # MNoP MAX Stat

    # ### Payment Seekers
    row = 4
    while row <= t_rows:
        PFinder = "=IF(Or($F" + str(row+1) + "=0,$F" + str(row+1) + '=""),"",Max(' + column[Edge2] + '$3:' + column[Edge2] + str(row) + ")+1)"
        PChecker = "=IF(" + column[Edge2] + str(row+1) + '="","",A' + str(row+1)+')'
        worksheet.write(row,Edge1,PFinder)  # If we find a payment, Max + 1
        worksheet.write(row,Edge2,PChecker)   # If the previous code found a payment, write it's line number
        row +=1

    # ######################### Payment Validator Columns #################################
    col = 20  # Col U
    while col <= spread_pad:
        # ### Row 1: Find If row number is <= Max, add if so
        ChkAMod = workbook.add_format({'bg_color':'#daeef3','align':'center'})
        Check_A = '=IF(' + column[col] + "1<$" + column[Edge2] + "1," + column[col] + '1+1,"")'
        #print(Check_A)
        worksheet.write(0,col,Check_A,ChkAMod)


        # ### Row 2: The Payment Row
        Pcol = column[col+1]
        ChkBMod = workbook.add_format({'bg_color':'#ffffcc','align':'center'})
        subsetB = [
            '=IF(' + Pcol + '1="",'
            '"",'
            'Vlookup(' + Pcol + '1,$' + column[Edge2] + '3:$' + column[Edge2+1] + str(t_rows) + ',2)'
            ')']
        Check_B = ""
        for line in subsetB:
            Check_B += line

        worksheet.write(1,col,Check_B,ChkBMod)
        
        # ### Row 3: Payment Amount | =Round(Indirect("F"&Col1) - sum(Col$3:Colrow-1) -sum(Colrow+1:Col$200),2)
        ChkCMod = workbook.add_format({'num_format': '[$$-409]#,##0.00','bg_color':'#f2f2f2','font_color':'#fa7d00','bold':True,'bottom':1})
        subsetC = [
            '=IfError('
            'Round(Indirect("F"&' + Pcol + '2+2)-',
            'Sum(' + Pcol + '$4:' + 'Indirect("' + Pcol +'"&(' + Pcol + '2+1)))' + '-',
            'Sum(Indirect("' + Pcol + '"&(' + Pcol + '2+3)):' + Pcol + '$250)'+ ',2)',
            ',"")'
            ]
        Check_C = ""
        for line in subsetC:
            Check_C += line
        worksheet.write(2,col,Check_C,ChkCMod)

        # ### Row 4-200: Double Check whether the full amount of payment is used or not.
        row = 3
        while row <= 250:
            crow = row + 1
            Codeline = ""
            subsets = [
                '=IF(' + Pcol + '$1="","",'
                'IF($A' + str(row+1) + '=' + Pcol + '$2,',
                'IF(' + Pcol + '$3=0,"Cleared",Concatenate(' + Pcol + '$2," :  ",' + Pcol + '$3)),',
                "IF($K" + str(crow) + "=" + Pcol + '$2,$L' + str(crow) + ',',
                "IF($M" + str(crow) + "=" + Pcol + '$2,$N' + str(crow) + ',',
                "IF($O" + str(crow) + "=" + Pcol + '$2,$P' + str(crow) + ',',
                "IF($Q" + str(crow) + "=" + Pcol + '$2,$R' + str(crow) + ',"")))))'
                ')'
                ]
            for line in subsets:
                Codeline += line
            worksheet.write(row,col,Codeline)
            row += 1
        col += 1

    # ### Total "I" Column conditional formatting to show where the balance zeroes out clearer.
    ZGreen = '#c6efce'
    ZeroQ = workbook.add_format({'num_format': '[$$-409]#,##0.00','bg_color': ZGreen})  # For Col I
    I_range = 'I3:I' + str(t_rows) # Initially 'I3:I200'
    worksheet.conditional_format(I_range,{'type':'cell','criteria':'between','minimum':-0.01,'maximum':0.01,'format':ZeroQ})

    # ### U4:Edge cells, check for "Cleared" and turn it green, otherwise, nada
    Cleared = workbook.add_format({'bg_color':ZGreen})  # For when a Cleared indicator shows up
    UEnd_range = 'U4:' + column[spread_pad] + str(t_rows)  # Initially 'U4:BZ250'
    worksheet.conditional_format(UEnd_range,{'type':'cell','criteria':'=','value':'"Cleared"','format':Cleared})
    
    worksheet.print_area(0, 0, t_rows, 17)
    worksheet.fit_to_pages(1, 0)
    workbook.close()
    return
# ################################### Ledger Sorter #######################################


def Ledger_Sorter(PRCN):  # Go through the Ledger List and pull anything associated with the PCRN
    global Ledger
    sm_ledger = []

    for line in Ledger:
        prcn = line[0]
        if prcn == PRCN:
            sm_ledger.append(line)
        else:
            pass
    output = reSort(sm_ledger)

    if len(output) >= 1:
        Ledger_Export(sm_ledger, PRCN)
    else:
        pass

# #################################### PRCN Lookup ########################################
for patient in Rollerdex:
    Ledger_Sorter(patient)