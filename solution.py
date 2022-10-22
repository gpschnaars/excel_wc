import pandas as pd

###########################################################
## extract target words and letter matrix from the excel ##
###########################################################
df = pd.read_excel(
    '0. Qualification Round. The Number 1 Fan - 5 PM Session.xlsx',
    sheet_name = 'Wally (L2)',
    skiprows = 1,
)

targets = df['Word to Find'].dropna().tolist()
matrix = df[df.columns[2:22]].values

####################
## solution loops ##
####################

results = dict()

# horizontal check 
for n_row, line in enumerate(matrix, 1):
    line_str = ''.join(line)
    for target in targets:
        if target not in results:
            if (target in line_str) or (target[::-1] in line_str):
                results[target] = n_row

# vertical check 
for line in matrix.transpose():
    line_str = ''.join(line)
    for target in targets:
        if target not in results:
            forward = line_str.find(target)
            if forward > -1:
                results[target] = forward + 1
                continue
            backward = line_str.find(target[::-1])
            if backward > -1:
                results[target] = backward + len(target)

###################################
## show result in original order ##
###################################
for target in targets:
    print(f"{target}: {results[target]}")

##
# WALLY: 1
# FMWC: 419
# MSFT: 533
# EXCEL: 973
# ESPORTS: 855
# LUCK: 465
# THINK: 98
# PLAN: 700
# COUNT: 379
# PREPARE: 541
# EXECUTE: 17
# STRONG: 851
# MEWC: 178
# FAST: 642
# SOLVE: 2
# NICE: 725
# GREAT: 706
# SPACE: 134
# LOGIC: 176
# FOCUS: 175
# FINISH: 134