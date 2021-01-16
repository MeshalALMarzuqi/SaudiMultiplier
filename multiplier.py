# Importing the required libaries
import pandas as pd
import numpy as np
import os


# checking the working directory
os.getcwd()
 

# Stating the file location
xls = pd.ExcelFile ('https://github.com/MeshalALMarzuqi/SaudiMultiplier/raw/main/SnU2016.xlsx')

# Reading the data in pandas DataFrame
df = pd.read_excel(xls, sheet_name = 'supply and use',na_values='n/a')
ldf = pd.read_excel(xls, sheet_name = 'labour', na_values='n/a')
ldf = pd.concat([ldf,(ldf['Total']-ldf['Saudi']).rename('Non Saudi')],
                axis=1, sort=False).fillna(0)
sdf = pd.read_excel(xls, sheet_name = 'survey', na_value='n/a' )
sdf = pd.concat([sdf,(sdf['Total']-sdf['Saudi']).rename('Non Saudi')]
                ,axis=1, sort=False).fillna(0)


# Assigning variables
TSPP = df.loc['Total Supply at Purchers Prices']
PC = df.loc['Private Consumption']
LC = df.loc['Labour Compensation']
VA = df.loc['Value Added']
tran_margines = df.loc['Transport margines']
trade_margines = df.loc['Trade margines']
TSPP = TSPP - tran_margines - trade_margines

industry = ['Agriculture','Fishing','Crude Petroleum & Natural Gas',
            'Other Mining','Petroleum Refining','Petrochemicals',
            'Other Manufacturing','Electricity gas and water supply',
            'Construction','Wholesale','Hotels and restaurants','Transport',
            'Financial intermediation','Real estate','Public administration',
            'Education','Health','Other community',
            'Private households with employed persons',
            'Services provided by extra-territorial organizations and bodies']




               ####### without induced effect #########

# Creating a use table 
usetable1 = df.drop(['Total Supply at Purchers Prices','Private Consumption',
                   'Labour Compensation','Value Added','Transport margines',
                    'Trade margines'])

# Creating A matrix 
A1 = usetable1 / TSPP
A1 = A1.fillna(0)

# Creating identity matix and converting it to a pandas dataframe
I1 = np.eye(len(usetable1))
I1 = pd.DataFrame(I1, index = usetable1.index, columns = usetable1.columns)

# identity matrix - A matrix 
I1_A1 = I1 - A1

# Checking if the determint is not equal to zero
#print(np.linalg.det(I1_A1))

# Checking if the determint is not equal to zero
#print(np.linalg.det(I1_A1))
# Calculating the inverse of I1_A1 and converting it to a pandas dataframe
Inv1 = np.linalg.inv(I1_A1)
Inv1 = pd.DataFrame(Inv1, index = usetable1.index, columns = usetable1.columns)
               
               ####### with induced effect #########
    
# Creating a use table & A matrix
usetable2 = df.drop(['Total Supply at Purchers Prices','Private Consumption',
                   'Value Added','Transport margines','Trade margines'])

# Creating A matrix
A2 = usetable2 / TSPP
PCratio = PC / (TSPP.sum() - usetable1.sum().sum()) # Note dividing by the sum 
                                                    # of TSPP minus the sum of
                                                    # the usetable1 
A2 = pd.concat([A2,PCratio.rename('Private Consumption')], axis =1,
              sort = False).fillna(0)
TSPP_USE1 = TSPP.sum() - usetable1.sum().sum()
A2.loc['Labour Compensation']['Private Consumption'] = PC.sum()/TSPP_USE1
usetable2 = pd.concat([usetable2,PC.rename('Private Consumption')], axis=1,
                      sort = False).fillna(0)
usetable2.loc['Labour Compensation','Private Consumption'] = PC.sum()

# Creating identity matix and converting it to a pandas dataframe 
I2 = np.eye(len(usetable2))
I2 = pd.DataFrame(I2, index = usetable2.index, columns = usetable2.columns)

# identity matrix - A matrix
I2_A2 = I2 - A2 

# Checking if the determint is not equal to zero
#print(np.linalg.det(I2_A2))

# Calculating the inverse of I1_A1 and converting it to a pandas dataframe
Inv2 = np.linalg.inv(I2_A2)
Inv2 = pd.DataFrame(Inv2, index = usetable2.index, columns = usetable2.columns)


# Cleaning Labour Data

# Calculating Labours' ratio in each economic activity
fish_ratio = sdf.iloc[2] / sdf.iloc[:3].sum()
agri_ratio = 1 - fish_ratio
crude_ratio = sdf.iloc[4] / sdf.iloc[3:8].sum()
other_min_ratio = 1 - crude_ratio
refine_ratio = sdf.iloc[17] / sdf.iloc[8:31].sum()
petro_ratio = sdf.iloc[18] / (sdf.iloc[8:31].sum() - sdf.iloc[17])
other_manu_ratio = 1 - refine_ratio - petro_ratio

# assembleing labours by economic activity
agri = ldf.iloc[0] * agri_ratio['Total']  
fish = ldf.iloc[0] * fish_ratio['Total']
crude = ldf.iloc[1] * crude_ratio['Total']
other_min = ldf.iloc[1] * other_min_ratio['Total']
refine = ldf.iloc[2] * refine_ratio['Total']
petro = ldf.iloc[2] * petro_ratio['Total']
other_manu = ldf.iloc[2] * other_manu_ratio['Total']
egw = ldf.iloc[3:5].sum()
cons = ldf.iloc[5]
whole = ldf.iloc[6]
acco = ldf.iloc[8]
trainf = ldf.iloc[7] + ldf.iloc[9]
fin = ldf.iloc[10]
real = ldf.iloc[11]
pradpu = ldf.iloc[12:15].sum()
edu = ldf.iloc[15]
heal = ldf.iloc[16]
othcom = ldf.iloc[17]+ldf.iloc[18]+ldf.iloc[20]
hold = ldf.iloc[19]
serv = ldf.iloc[21]

# Creating Jobs Data Frame
LIST = [agri, fish, crude, other_min, refine, petro, other_manu, egw, cons, 
        whole, acco, trainf, fin, real, pradpu, edu, heal, othcom, hold, serv]
Jobs = pd.concat(LIST, axis = 1)
Jobs.columns = industry

# Total Jobs
TJ = Jobs.loc['Total']

# Saudi Jobs
SJ = Jobs.loc['Saudi']

# Non Saudi Jobs
NSJ = TJ - SJ

# Transposing Jobs DataFrame
Jobs_T = Jobs.T


# Calculating Output Multipliers:

# Total Output Multiplier With Induced Effect
TOMI = Inv2.sum()
TOMI = TOMI[TSPP.index]

# Total Output Multiplier Without Induced Effect
TOM = Inv1.sum()

# Direct Output Multiplier
DOM = A1.sum()

# Indirect Output Multiplier
IOM = TOM - 1 - DOM

# Induced Output Multiplier
IndOM = (TOMI - TOM)
IndOM = IndOM[TSPP.index]

# Type 1 and Type 2 Output Multipliers
t1_output = (DOM + IOM) / DOM
t2_output = (DOM + IOM + IndOM) / DOM


# Calculating GDP Multipliers:

# Calculating the ratio of Valude Added to Total Supply at Purchers Prices 
VAr = VA / TSPP

# Total GDP Multiplier With Induced Effect
TGMI = Inv2.mul(VAr, axis = 0).sum()
TGMI = TGMI[TSPP.index]

# Total GDP Multiplier Without Induced Effect
TGM = Inv1.mul(VAr, axis = 0).sum()

# Direct GDP Mutiplier
DGM = A1.mul(VAr, axis = 0).sum()

# Indirect GDP Multiplier
IGM = TGM - VAr - DGM

# Induced Jobs Multiplier
IndGM = TGMI - TGM
IndGM = IndGM[TSPP.index]

# Type 1 and Type 2 GDP Multipliers
t1_gdp = (DGM + IGM) / DGM
t2_gdp = (DGM + IGM + IndGM) / DGM


# Calculating Jobs Multipliers:

# Calculating the ratio of Total Jobs to Total Supply at Purchers Prices 
Jr = TJ / TSPP

# Total Jobs Multiplier With Induced Effect
TJMI = Inv2.mul(Jr, axis = 0).sum()
TJMI = TJMI[TSPP.index]

# Total Jobs Multiplier Without Induced Effect
TJM = Inv1.mul(Jr, axis = 0).sum()

# Direct Jobs Mutiplier
DJM = A1.mul(Jr, axis = 0).sum()

# Indirect Jobs Multiplier
IJM = TJM - Jr - DJM
IJM = IJM[TSPP.index]

# Induced Jobs Multiplier
IndJM = TJMI - TJM
IndJM = IndJM[TSPP.index]

# Type 1 and Type 2 Jobs Multipliers
t1_job = (DJM + IJM) / DJM
t1_job = t1_job[TSPP.index]
t2_job = (DJM + IJM + IndJM) / DJM
t2_job = t2_job[TSPP.index]


# Printing Data to Excel 


MultiIndex = ['Total Multiplier with Induced Effect', 
              'Total Multiplier without Induced Effect', 'Direct Multiplier',
             'Indirect Multiplier', 'Induced Multiplier','Type 1 Multiplier',
             'Type 2 Multiplier']

# Create Pandas DataFrames
OutPutMulti = pd.concat([TOMI, TOM, DOM, IOM, IndOM, t1_output, t2_output], 
                        axis=1, sort = False)
OutPutMulti.columns = MultiIndex

GDPMulti = pd.concat([TGMI, TGM, DGM, IGM, IndGM, t1_gdp, t2_gdp], 
                     axis=1, sort = False)
GDPMulti.columns = MultiIndex

JobsMulti = pd.concat([TJMI, TJM, DJM, IJM, IndJM, t1_job, t2_job], 
                      axis=1, sort = False)
JobsMulti.columns = MultiIndex

# Write each dataframe to a different worksheet.
writer = pd.ExcelWriter('2016Multipliers.xlsx', engine='xlsxwriter')
OutPutMulti.to_excel(writer, sheet_name='Output Multiplier')
GDPMulti.to_excel(writer, sheet_name='GDP Multiplier')
JobsMulti.to_excel(writer, sheet_name='Jobs Multiplier')
Jobs_T.to_excel(writer, sheet_name='Labour')

# Close the Pandas Excel writer and output the Excel file.
writer.save()