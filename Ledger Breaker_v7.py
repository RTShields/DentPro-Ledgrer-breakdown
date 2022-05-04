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

if os.path.isdir('/Ledgers'):
    pass
else:
    os.mkdir('/Ledgers')

# ################################# Patient Splicer ###############################
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


# ################################ Row Restructuring ##############################
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
        print(item03)
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


# ################################### Sort Ledger ################################


def reSort(ledger):
    # new_list = sorted(a_list, key=lambda x: (len(x), x))
    # https://www.kite.com/python/answers/how-to-sort-by-two-keys-in-python
    sorter = sorted(ledger, key=lambda x: (x[1],x[2]))
    return sorter

# ################################# Ledger Rebuild ###############################


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
    merge = workbook.add_format({'size': 20})
    title = workbook.add_format({'size': 20, 'bold': 1, 'align': 'center', 'valign': 'center'})
    bold = workbook.add_format({'bold': True})
    wingding = workbook.add_format({'font':'Wingdings','align':'center'})
    # ### Normal Rows
    accounting = workbook.add_format({'num_format': '[$$-409]#,##0.00'})
    money = workbook.add_format({'num_format': '$#,##'})
    date = workbook.add_format({'num_format': 'mm/dd/yy','align':'center'})
    counting = workbook.add_format({'align': 'center'})
    blue = '#DAEEF3'
    green = '#EBF1DE'
    purple = '#e4dfec'
    # ### Insurance Rows
    ins_wing = workbook.add_format({'font': 'Wingdings','align':'center','bg_color':green}) 
    ins_date = workbook.add_format({'num_format': 'mm/dd/yy','align':'center','bg_color':green})
    ins_norm = workbook.add_format({'bg_color':green})
    ins_fund = workbook.add_format({'num_format': '[$$-409]#,##0.00','bg_color':green})
    # ### Payment Rows
    pay_wing = workbook.add_format({'font': 'Wingdings','align':'center','bg_color':blue})
    pay_date = workbook.add_format({'num_format': 'mm/dd/yy','align':'center','bg_color':blue})
    pay_norm = workbook.add_format({'bg_color':blue})
    pay_fund = workbook.add_format({'num_format': '[$$-409]#,##0.00','bg_color':blue})
    # ### Transfer Rows
    trf_wing = workbook.add_format({'font': 'Wingdings','align':'center','bg_color':purple})
    trf_date = workbook.add_format({'num_format': 'mm/dd/yy','align':'center','bg_color':purple})
    trf_norm = workbook.add_format({'bg_color':purple})
    trf_fund = workbook.add_format({'num_format': '[$$-409]#,##0.00','bg_color':purple})

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
    worksheet.write(2,0,1,counting)
    worksheet.write(2,1,'=B4',date)
    worksheet.write(2,3,'Starting Balance')
    worksheet.write(2,8,0.00,accounting)

    # ### Insert each ledger row
    payrow = []
    for row in range(len(ledger)):
        line = ledger[row]
        R = row + 3
        L_Date = line[1]
        L_Code = line[2]
        L_Desc = line[3]
        L_Wing = line[4]
        L_Cred = line[5]
        L_Debt = line[6]

        if line[5] == 0:
            worksheet.write(R,1,L_Date,date) # Date
            worksheet.write(R,2,L_Code) # ADA Code
            worksheet.write(R,3,L_Desc) # Line Description
            worksheet.write(R,4,L_Wing,wingding) # Line Wingding
            worksheet.write(R,6,L_Debt,money) # Debit
        else:
            payrow.append(R)
            PMT_types = 'o¡' # Insurance, Credits <- Catchall for everything except transfers
            if line[4] == PMT_types[0]:
                worksheet.write(R,1,L_Date,ins_date)    # Date
                worksheet.write(R,2,"",ins_norm)        # Code
                worksheet.write(R,3,L_Desc,ins_norm)    # Line Description
                worksheet.write(R,4,L_Wing,ins_wing)    # Wingding
                worksheet.write(R,5,L_Cred,ins_fund)    # Credit
                worksheet.write(R,6,"",ins_norm)        # Debit
                worksheet.write(R,7,"",ins_norm)        # Space
            elif line[4] == PMT_types[1]:
                worksheet.write(R,1,L_Date,pay_date)
                worksheet.write(R,2,"",pay_norm)
                worksheet.write(R,3,L_Desc,pay_norm)
                worksheet.write(R,4,L_Wing,pay_wing)
                worksheet.write(R,5,L_Cred,pay_fund)
                worksheet.write(R,6,"",pay_norm)
                worksheet.write(R,7,"",pay_norm)
            else:
                pass
        ldesc = L_Desc.lower()
        if ldesc.find('transfer') != -1:
            worksheet.write(R,1,L_Date,trf_date)
            worksheet.write(R,2,"",trf_norm)
            worksheet.write(R,3,L_Desc,trf_norm)
            worksheet.write(R,4,'²',trf_wing)
            if ldesc.find('to') != -1:
                worksheet.write(R,5,"",trf_norm)
                worksheet.write(R,6,L_Debt,trf_fund)
            else:
                worksheet.write(R,5,L_Cred,trf_fund)
                worksheet.write(R,6,"",trf_norm)
            worksheet.write(R,7,"",trf_norm)

    # ### Column Coding
    row = 3
    counter = 2 + len(ledger)
    while row < 201:
        R = str(row)
        R1 = str(row+1)
        codeA = '=IF(D' + R1 + '="","",Sum(Max(A$3:A' + R + ')+1))'
        codeB = '=IF(D' + R1 + '="","",B' + R + ')'
        codeI = '=IF(A' + R1 + '="","",I' + R + '+(F' + R1 + '-G' + R1 + '))'
        codeS = '=IF(C' + R1 + '="","",G' + R1 + '-Sum(L' + R1 + ',N' + R1 + ',P' + R1  + ',R' + R1 + '))'
        worksheet.write(row,0,codeA,counting)
        if row > counter:
            worksheet.write(row,1,codeB,date)
        worksheet.write(row,8,codeI,accounting)
        worksheet.write(row,18,codeS,accounting)

        # ### Border walls around payment section
        wall_L = workbook.add_format({'left':1})
        wall_R = workbook.add_format({'num_format': '[$$-409]#,##0.00'})
        worksheet.write(row,10,"",wall_L)
        worksheet.write(row,11,"",wall_R)
        worksheet.write(row,12,"",wall_L)
        worksheet.write(row,13,"",wall_R)
        worksheet.write(row,14,"",wall_L)
        worksheet.write(row,15,"",wall_R)
        worksheet.write(row,16,"",wall_L)
        worksheet.write(row,17,"",wall_R)

        row += 1

    # ### Find and assign payments to CA and CB
    row = 4
    worksheet.write(3,78,0)  # Counter start
    worksheet.write(0,78,'=Max(CA3:CA250)')  # Max number of payments
    while row <= 200:
        c_code = 78
        # If the F# cell has a payment, assign it a number
        worksheet.write(row,c_code,"=IF(Or($F" + str(row+1) + "=0,$F" + str(row+1) + '=""),"",Max(CA$3:CA' + str(row) + ")+1)")
        # If CA# has a number, grab it's row #
        c_code +=1
        worksheet.write(row,c_code,"=IF(CA" + str(row+1) + '="","",A' + str(row+1)+')')
        row +=1

    # ### Create payment columns for U - BZ using CA/CB
    col = 20  # Col U
    column = (0,'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AY','BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY','BY','BZ')

    while col <= 77:
        # ### Row 1: Find If row number is <= Max, add if so
        ChkAMod = workbook.add_format({'bg_color':'#daeef3','align':'center'})
        Check_A = '=IF(' + column[col] + "1<$CA1," + column[col] + '1+1,"")'
        worksheet.write(0,col,Check_A,ChkAMod)

        # ### Row 2: The Payment Row
        Pcol = column[col+1]
        ChkBMod = workbook.add_format({'bg_color':'#ffffcc','align':'center'})
        Check_B = '=Vlookup(' + Pcol + '1,$CA$3:$CB250,2)'
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
        while row <= 200:
            crow = row + 1
            Codeline = ""
            subsets = [
                '=IF(' + Pcol + '1="","",'
                'IF(A' + str(row+1) + '=' + Pcol + '2,',
                'IF(' + Pcol + '3=0,"Cleared",Concatenate(' + Pcol + '2," :  ",' + Pcol + '3)),',
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
    worksheet.conditional_format('I3:I201',{'type':'cell','criteria':'=','value':0,'format':ZeroQ})

    # ### U4:BZ250 cells, check for "Cleared" and turn it green, otherwise, nada
    Cleared = workbook.add_format({'bg_color':ZGreen})  # For 
    worksheet.conditional_format('U4:BZ250',{'type':'cell','criteria':'=','value':'"Cleared"','format':Cleared})

    workbook.close()
    return
# ################################# Ledger Sorter  ###############################


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

# ################################### PRCN Lookup ################################

for patient in Rollerdex:
    Ledger_Sorter(patient)