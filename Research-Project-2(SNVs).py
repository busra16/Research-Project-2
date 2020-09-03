
#openpyxl is a library that provide read excel files.
import openpyxl
#XlsxWriter is a Python module for creating Excel XLSX files.
import xlsxwriter
#It is an extension of the Numpy library. Used where Numpy is inadequate.
#Because in Numpy, the rows and columns of the array must be homogeneous.
#The data structure is series and dataframes.
import pandas as pd
#Enables data to be read from excel
import xlrd

output_file = "result.xlsx"
input_excels = ["GRCh38 - G05830_Proband_LP3000194-DNA_C10_LP3000194-DNA_C10 (1).xlsx",
                "GRCh38 - G100285_Proband_LP3000598-DNA_D12_LP3000598-DNA_D12.xlsx",
                "GRCh38 - G100323_Proband_LP3000629-DNA_B08_LP3000629-DNA_B08.xlsx",
                "GRCh38 - G100519_Proband_LP3000177-DNA_B06_LP3000177-DNA_B06.xlsx",
                "GRCh38 - G101570_Proband_LP3000836-DNA_F03_LP3000836-DNA_F03.xlsx",
                "GRCh38 - G101612_Proband_LP3000458-DNA_A10_LP3000458-DNA_A10.xlsx",
                "GRCh38 - G103245_Proband_LP3000598-DNA_F12_LP3000598-DNA_F12.xlsx",
                "GRCh38 - G103660_Proband_LP3000292-DNA_E12_LP3000292-DNA_E12.xlsx",
                "GRCh38 - G104026_Proband_LP3001043-DNA_C11_LP3001043-DNA_C11.xlsx",
                "GRCh38 - G106418_Proband_LP3000276-DNA_B06_LP3000276-DNA_B06.xlsx",
                "GRCh38 - G106537_Proband_LP3000292-DNA_F04_LP3000292-DNA_F04.xlsx",
                "GRCh38 - G113993_Proband_LP3001190-DNA_D01_LP3001190-DNA_D01.xlsx",
                "GRCh38 - G31037_Proband_LP3000118-DNA_E01_LP3000118-DNA_E01.xlsx",
                "GRCh38 - G39752_Proband_LP3000598-DNA_F03_LP3000598-DNA_F03.xlsx",
                "GRCh38 - G50966_Proband_LP3000122-DNA_E11_LP3000122-DNA_E11.xlsx",
                "GRCh38 - G62347_Proband_LP3000629-DNA_F11_LP3000629-DNA_F11.xlsx",
                "GRCh38 - G71665_Proband_LP3001079-DNA_A09_LP3001079-DNA_A09.xlsx",
                "GRCh38 - G74543_Proband_LP3000323-DNA_E02_LP3000323-DNA_E02.xlsx",
                "GRCh38 - G76448_Proband_LP3000317-DNA_D04_LP3000317-DNA_D04.xlsx",
                "GRCh38 - G76499_Proband_LP3000186-DNA_A06_LP3000186-DNA_A06.xlsx",
                "GRCh38 - G79629_Proband_LP3000256-DNA_A06_LP3000256-DNA_A06.xlsx",
                "GRCh38 - G80057_Proband_LP3000658-DNA_F06_LP3000658-DNA_F06.xlsx",
                "GRCh38 - G80251_Proband_LP3000185-DNA_F04_LP3000185-DNA_F04.xlsx",
                "GRCh38 - G81145_Proband_LP3000310-DNA_F05_LP3000310-DNA_F05.xlsx",
                "GRCh38 - G82498_Proband_LP3000230-DNA_D09_LP3000230-DNA_D09.xlsx",
                "GRCh38 - G82899_Proband_LP3000102-DNA_H04_LP3000102-DNA_H04.xlsx",
                "GRCh38 - G87579_Proband_LP3000629-DNA_A06_LP3000629-DNA_A06.xlsx",
                "GRCh38 - G90528_Proband_LP3000655-DNA_G02_LP3000655-DNA_G02.xlsx",
                "GRCh38 - G90546_Proband_LP3000466-DNA_G10_LP3000466-DNA_G10.xlsx",
                "GRCh38 - G94496_Proband_LP3000256-DNA_F04_LP3000256-DNA_F04.xlsx",
                "GRCh38 - G97804_Proband_LP3000256-DNA_D03_LP3000256-DNA_D03.xlsx",
                "GRCh38 - G98965_Proband_LP3000750-DNA_H01_LP3000750-DNA_H01.xlsx"]
input_beds = ["promoters/hg19_droppped_CAGE_peaks_phase1and2.bed",
                "promoters/hg38_fair+new_CAGE_peaks_phase1and2.bed",
                "promoters/hg38_fair_CAGE_peaks_phase1and2.bed",
                "promoters/hg38_liftover+new_CAGE_peaks_phase1and2.bed",
                "promoters/hg38_liftover_CAGE_peaks_phase1and2.bed",
                "promoters/hg38_new_CAGE_peaks_phase1and2.bed",
                "promoters/hg38_problematic_CAGE_peaks_phase1and2.bed"]

result = [["Chr:Pos", "Ref/Alt", "Promoter start", "Promoter end"]]
keys = ['chr1', 'chr2', 'chr3', 'chr4', 'chr5', 'chr6', 'chr7', 'chr8', 'chr9',
        'chr10', 'chr11', 'chr12', 'chr13', 'chr14', 'chr15', 'chr16', 'chr17',
        'chr18', 'chr19', 'chr20', 'chr21', 'chr22', 'chrX', 'chrY']

#-------------------------------------------------------------------------------------------
#separates the chromosome positions in both excel files of the patient.
chrs = []
initial_promoters = []
promoter_ranges = []

def get_chromsomes(file):
    chrs.append([file, ''])
    #read the file
    wb = xlrd.open_workbook(file)
     # get the first worksheet
    s = wb.sheet_by_index(0)

# adds the data in each row in the zeroth column to the chrs list.
    for row in range(2, s.nrows):
        chr_ref = []
        for col in range(0, 2):
            chr_ref.append(s.cell(row,col).value)
        chrs.append(chr_ref)

    return chrs
#With this defined function, chromosome positions in the patient file were assigned to the list.
    
#------------------------------------------------------------------------------------------
##Promoter files are listed. And it deletes the spaces at the
# beginning and end of the file.
def get_promoter_ranges(file):

    with open(file)as f:
        for line in f:
            initial_promoters.append(line.strip().split())

    for initial_promoter in initial_promoters:
        line = []
        line.append(initial_promoter[0])
        line.append(initial_promoter[1])
        line.append(initial_promoter[2])
        promoter_ranges.append(line)

def compare(chroms, pros):
    pros = list(set(map(lambda i: tuple(sorted(i)), pros)))

    for chrom in chroms:
        value = chrom[0].split(':')[1]
        for pro in pros:
            line = []
            if int(value) > int(pro[0]) and int(value) < int(pro[1]):
                line.append(chrom[0])
                line.append(chrom[1])
                line.append(pro[0])
                line.append(pro[1])
                result.append(line)

def outputExcel(res):
    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet()
    for i, l in enumerate(res):
        for j, col in enumerate(l):
            worksheet.write(i, j, col)

    workbook.close()

for input_bed in input_beds:
    get_promoter_ranges(input_bed)

for input_excel in input_excels:
    chromsomes = get_chromsomes(input_excel)
    result.append([input_excel, "", "", ""])
    for key in keys:
        compare_promoter_ranges = []
        compare_chromsomes = []

        for promoter_range in promoter_ranges:
            promoter_line = []
            if promoter_range[0] == key:
                promoter_line.append(promoter_range[1])
                promoter_line.append(promoter_range[2])
                compare_promoter_ranges.append(promoter_line)

        for chromsome in chromsomes:
            if 'chr' + chromsome[0].split(':')[0] == key:
                compare_chromsomes.append(chromsome)

        compare(compare_chromsomes, compare_promoter_ranges)
        
outputExcel(result)