"""
Created on Tue Oct 20 12:38:10 2020

@author: bio608
"""
import openpyxl
import re
import csv
listFeatures=["246","247","249","264","265","268","250","257","262","271","273","286","288","289","291","306","307","310","292","299","304","313","315","328","490","491","493","506","507","510","494","499","504","513","515","528","530","531","533","548","549","552","534","541","546","555","557","570","572","573","575","589","590","593","576","582","587","596","598","611","613","614","616","630","631","634","617","623","628","637","639","652","723","724","726","741","742","745","727","734","739","748","750","763","765","766","768","783","784","787","769","776","781","790","792","805","807","808","810","823","824","827","811","816","821","830","832","845","910","911","913","928","929","932","914","921","926","935","937","950","1028","1029","1031","1046","1047","1050","1032","1039","1044","1053","1055","1068","1070","1071","1073","1088","1089","1092","1074","1081","1086","1095","1097","1111","1113","1114","1116","1131","1132","1135","1117","1124","1129","1138","1140","1153","1155","1156","1158","1173","1174","1177","1159","1166","1171","1180","1182","1195","1197","1198","1200","1214","1215","1218","1201","1208","1212","1221","1223","1236","1238","1239","1241","1256","1257","1260","1242","1249","1254","1263","1265","1278","1280","1281","1283","1298","1299","1302","1284","1291","1296","1305","1307","1320","1322","1323","1325","1341","1342","1345","1326","1333","1339","1348","1350","1361","1363","1364","1366","1382","1383","1386","1367","1374","1380","1389","1391","1402","1404","1405","1407","1422","1423","1426","1408","1415","1420","1429","1431","1444","1446","1447","1449","1463","1464","1467","1450","1456","1461","1470","1472","1485","1487","1488","1490","1504","1505","1508","1491","1497","1502","1511","1513","1526","1606","1607","1609","1624","1625","1628","1610","1617","1622","1631","1633","1646","1648","1649","1651","1665","1666","1669","1652","1658","1663","1672","1674","1687","1766","1767","1769","1784","1785","1788","1770","1777","1782","1791","1793","1806","1808","1809","1811","1826","1827","1830","1812","1818","1824","1833","1835","1847","1849","1850","1852","1867","1868","1871","1853","1860","1865","1874","1876","1889","1891","1892","1894","1907","1908","1911","1895","1900","1905","1914","1916","1928","1930","1931","1933","1949","1950","1953","1934","1942","1947","1956","1958","1971","1973","1974","1976","1991","1992","1995","1977","1984","1989","1998","2000","2011","2013","2014","2016","2031","2032","2035","2017","2024","2029","2038","2040","2053","2127","2128","2130","2144","2145","2148","2131","2137","2142","2151","2153","2164","2166","2167","2169","2183","2184","2187","2170","2176","2181","2190","2192","2204","2206","2207","2209","2223","2224","2227","2210","2217","2221","2230","2232","2245","2247","2248","2250","2267","2268","2271","2251","2258","2265","2274","2276","2289","2291","2292","2294","2309","2310","2313","2295","2302","2307","2316","2318","2331","2402","2403","2405","2420","2421","2424","2406","2413","2418","2427","2429","2442","2444","2445","2447","2461","2462","2465","2448","2455","2459","2468","2470","2483","2485","2486","2488","2502","2503","2506","2489","2496","2500","2509","2511","2522","2524","2525","2527","2543","2544","2547","2528","2535","2541","2550","2552","2565","2567","2568","2570","2583","2584","2587","2571","2577","2581","2590","2592","2605","2607","2608","2610","2624","2625","2628","2611","2617","2622","2631","2633","2646","2648","2649","2651","2665","2666","2669","2652","2659","2663","2672","2674","2686"]

#fnxl2 = 'C:/Users/bio608/Desktop/邱鈺凱/Spyder Marfan/exon+domain.xlsx'
#wb2 = openpyxl.load_workbook(fnxl2)
#ws2 = wb2.get_sheet_by_name('sheet2')
#lsexon=list(ws2.columns)[10][1:68]
#lsdomain=list(ws2.columns)[11][1:68]

csvfn = open('C:/Users/bio608/Desktop/邱鈺凱/New Marfan/MFS.csv','w', newline = '')
csvWrite = csv.writer(csvfn)#
csvWrite.writerow(['Location','Nucleotide','Protein','Domain','key residues','Mutation Type','Effect'])

fncsv1 = open('C:/Users/bio608/Desktop/邱鈺凱/Spyder Marfan/exon.csv')
csvreader1 = csv.reader(fncsv1)
listreport1 = list(csvreader1)[1:]

fncsv2 = open('C:/Users/bio608/Desktop/邱鈺凱/Spyder Marfan/domain.csv')
csvreader2 = csv.reader(fncsv2)
listreport2 = list(csvreader2)[1:] 

bases = "TCAG"
codons = [a + b + c for a in bases for b in bases for c in bases]
amino_acids = 'FFLLSSSSYY**CC*WLLLLPPPPHHQQRRRRIIIMTTTTNNKKSSRRVVVVAAAADDEEGGGG'
amino_table = {'A':'Ala','R':'Arg','N':'Asn','D':'Asp','*':'*','B':'Asx','C':'Cys','Q':'Gln','E':'Glu','Z':'Glx','G':'Gly','H':'His','I':'Ile','L':'Leu','K':'Lys','M':'Met','F':'Phe','P':'Pro','S':'Ser','T':'Thr','W':'Trp','Y':'Try','V':'Val'}
codon_table = dict(zip(codons, amino_acids))
complementary_table = {'A':'T','T': 'A', 'G':'C', 'C':'G'}
c_seq = ""
    
fn = open('C:/Users/bio608/Desktop/邱鈺凱/New Marfan/FBN1 mRNA  CDS(8616)(2).txt')
seq = ""
for nt in fn.read():
    if nt != ' ' and nt != '\n' and nt != '\t':
        seq += nt

fnxl = 'C:/Users/bio608/Desktop/邱鈺凱/New Marfan/MFS new list 改.xlsx'
wb = openpyxl.load_workbook(fnxl)
ws = wb.get_sheet_by_name('工作表2')

def Exon():
    for i in listreport1:
       if int(i[1])<=int(l[0])<=int(i[2]):
           ls.append(i[0])
       else:
           continue
def Intronplus():#c.G4816+1A
    for i in listreport1:
        if int(i[1])<=int(l[0])<=int(i[2]):
            I = re.sub("\D", "", i[0])
            ls.append("Intron{}".format(I))
        else:
            continue
def Intronminus():#c.G4816-1A
    for i in listreport1:
        if int(i[1])<=int(l[0])<=int(i[2]):
            I = re.sub("\D", "", i[0])
            ls.append("Intron{}".format(int(I)-1))
        else:
            continue
def Domain(R):#R餘數進位
    for j in listreport2:
        if int(j[1])<=int(int(l[0])/3)+int(R)<=int(j[2]):
            ls.append(j[0])
        else:
            continue

for cell in list(ws.columns)[3][1:171]:
    if cell.value != None:
        l=re.findall(r"\d+\.?\d*",str(cell.value))#cell.value取數字
        mute = ''.join(re.findall(r'[A-Za-z]', cell.value))#取字母
        ls=[]#寫入CSV檔
        
        if "d" in list(cell.value): 
            if "p" in list(cell.value):
                Exon()#def Exon()
                ls.append(cell.value)
                if int(l[0])% 3 == 0:
                    st=(str(seq[(int(l[0])-1)-2]),str(seq[(int(l[0])-1)-1]),str(seq[int(l[0])-1]))
                    code=""
                    cod=code.join(st)
                    aa = codon_table[cod]
                    ls.append("p.{}{}fs".format(aa,int(int(l[0])/3)))
                    Domain(0)#R餘數進位
                    ls.append('/t')
                    ls.append('Frameshift')
                    ls.append('Small duplication')
                    print(ls)
                    csvWrite.writerow([ls])
                elif int(l[0])% 3 == 1:
                    st=(str(seq[(int(l[0])-1)]),str(seq[(int(l[0])-1)+1]),str(seq[(int(l[0])-1)+2]))
                    code=""
                    cod=code.join(st)
                    aa = codon_table[cod]
                    ls.append("p.{}{}fs".format(aa,int(int(l[0])/3)+1))
                    Domain(1)#R餘數進位
                    ls.append('/t')
                    ls.append('Frameshift')
                    ls.append('Small duplication')
                    print(ls)
                    csvWrite.writerow([ls])
                elif int(l[0])% 3 == 2:
                    st=(str(seq[(int(l[0])-1)-1]),str(seq[(int(l[0])-1)]),str(seq[(int(l[0])-1)+1]))
                    code=""
                    cod=code.join(st)
                    aa = codon_table[cod]
                    ls.append("p.{}{}fs".format(aa,int(int(l[0])/3)+1))
                    Domain(1)#R餘數進位
                    ls.append('/t')
                    ls.append('Frameshift')
                    ls.append('Small duplication')
                    print(ls)
                    csvWrite.writerow([ls])
                else :
                    continue
            elif "s" in list(cell.value):
                Exon()#def Exon()
                ls.append(cell.value)
                
                if int(l[0]) % 3 == 0:
                    st=(str(seq[(int(l[0])-1)-2]),str(seq[(int(l[0])-1)-1]),str(seq[int(l[0])-1]))
                    code=""
                    cod=code.join(st)
                    aa = codon_table[cod]
                    ls.append("p.{}{}fs".format(aa,int(int(l[0])/3)))
                    Domain(0)#R餘數進位
                    ls.append('/t')
                    ls.append('Frameshift')
                    ls.append('Small delition/insertion')
                    print(ls)
                    csvWrite.writerow([ls])
                elif int(l[0]) % 3 == 1:
                    st=(str(seq[(int(l[0])-1)]),str(seq[(int(l[0])-1)+1]),str(seq[(int(l[0])-1)+2]))
                    code=""
                    cod=code.join(st)
                    aa = codon_table[cod]
                    ls.append("p.{}{}fs".format(aa,int(int(l[0])/3)+1))
                    Domain(1)#R餘數進位
                    ls.append('/t')
                    ls.append('Frameshift')
                    ls.append('Small delition/insertion')
                    print(ls)
                    csvWrite.writerow([ls])
                elif int(l[0]) % 3 == 2:
                    st=(str(seq[(int(l[0])-1)-1]),str(seq[(int(l[0])-1)]),str(seq[(int(l[0])-1)+1]))
                    code=""
                    cod=code.join(st)
                    aa = codon_table[cod]
                    ls.append("p.{}{}fs".format(aa,int(int(l[0])/3)+1))
                    Domain(1)#R餘數進位
                    ls.append('/t')
                    ls.append('Frameshift')
                    ls.append('Small delition/insertion')
                    print(ls)
                    csvWrite.writerow([ls])
          
            elif "+" in list(cell.value):
                Intronplus()#c.G4816+1A
                ls.append(cell.value)
                if int(l[1])<=2:
                    if int(l[0])% 3 == 0:
                        ls.append('/t')
                        Domain(0)#R餘數進位
                        ls.append('/t')
                        ls.append('Splice')
                        ls.append('Transcription')
                        print(ls)
                        csvWrite.writerow([ls])
                    elif int(l[0])% 3 == 1:
                        ls.append('/t')
                        Domain(1)#R餘數進位
                        ls.append('/t')
                        ls.append('Splice')
                        ls.append('Transcription')
                        print(ls)
                        csvWrite.writerow([ls])
                    elif int(l[0])% 3 == 2:
                        ls.append('/t')
                        Domain(1)#R餘數進位
                        ls.append('/t')
                        ls.append('Splice')
                        ls.append('Transcription')
                        print(ls)
                        csvWrite.writerow([ls])
                    else :
                        continue
                elif 2<int(l[1])<=10:
                    if int(l[0])% 3 == 0:
                        ls.append('/t')
                        Domain(0)#R餘數進位
                        ls.append('/t')
                        ls.append('Splice region')
                        ls.append('Transcription')
                        print(ls)
                        csvWrite.writerow([ls])
                    elif int(l[0])% 3 == 1:
                        ls.append('/t')
                        Domain(1)#R餘數進位
                        ls.append('/t')
                        ls.append('Splice region')
                        ls.append('Transcription')
                        print(ls)
                        csvWrite.writerow([ls])
                    elif int(l[0])% 3 == 2:
                        ls.append('/t')
                        Domain(1)#R餘數進位
                        ls.append('/t')
                        ls.append('Splice region')
                        ls.append('Transcription')
                        print(ls)
                        csvWrite.writerow([ls])
                elif 10<int(l[1]):
                    if int(l[0])% 3 == 0:
                        ls.append('/t')
                        Domain(0)#R餘數進位
                        ls.append('/t')
                        ls.append('Substitution')
                        ls.append('Transcription')
                        print(ls)
                        csvWrite.writerow([ls])
                    elif int(l[0])% 3 == 1:
                        ls.append('/t')
                        Domain(1)#R餘數進位
                        ls.append('/t')
                        ls.append('Substitution')
                        ls.append('Transcription')
                        print(ls)
                        csvWrite.writerow([ls])
                    elif int(l[0])% 3 == 2:
                        ls.append('/t')
                        Domain(1)#R餘數進位
                        ls.append('/t')
                        ls.append('Substitution')
                        ls.append('Transcription')
                        print(ls)
                        csvWrite.writerow([ls])
                else:
                    print('error')
                    
            else:
                Exon()#def Exon()
                ls.append(cell.value)
                if int(l[0])% 3 == 0:
                    if "_" in list(cell.value):
                        if int(int(l[1])-int(l[0]))<=10:
                            st=(str(seq[(int(l[0])-1)-2]),str(seq[(int(l[0])-1)-1]),str(seq[int(l[0])-1]))
                            code=""
                            cod=code.join(st)
                            aa = codon_table[cod]
                            ls.append("p.{}{}fs".format(aa,int(int(l[0])/3)))
                            Domain(0)#R餘數進位
                            ls.append('/t')
                            ls.append('Frameshift')
                            ls.append('Small deletion')
                            print(ls)
                            csvWrite.writerow([ls])
                        else:
                            st=(str(seq[(int(l[0])-1)-2]),str(seq[(int(l[0])-1)-1]),str(seq[int(l[0])-1]))
                            st2=(str(seq[(int(l[1])-1)-2]),str(seq[(int(l[1])-1)-1]),str(seq[int(l[1])-1]))
                            code=""
                            code2=""
                            cod=code.join(st)
                            cod2=code2.join(st2)
                            aa = codon_table[cod]
                            aa2 = codon_table[cod2]
                            ls.append("p.{}{}_{}{}del".format(aa,int(int(l[0])/3),aa2,int(int(l[1])/3)))
                            for j in listreport2:
                               if int(j[1])<=int((int(l[0])+int(l[1]))/2/3)<=int(j[2]):
                                   ls.append(j[0])
                                   
                            ls.append('/t')
                            ls.append('InFrame deletion')
                            ls.append('Gross deletion')
                            print(ls)
                            csvWrite.writerow([ls])
                            
                    else:
                        st=(str(seq[(int(l[0])-1)-2]),str(seq[(int(l[0])-1)-1]),str(seq[int(l[0])-1]))
                        code=""
                        cod=code.join(st)
                        aa = codon_table[cod]
                        ls.append("p.{}{}fs".format(aa,int(int(l[0])/3)))
                        Domain(0)#R餘數進位
                        ls.append('/t')
                        ls.append('Frameshift')
                        ls.append('Small deletion')
                        print(ls)
                        csvWrite.writerow([ls])
                elif int(l[0])% 3 == 1:
                    if "_" in list(cell.value):
                        if int(int(l[1])-int(l[0]))<=10:
                            st=(str(seq[(int(l[0])-1)]),str(seq[(int(l[0])-1)+1]),str(seq[(int(l[0])-1)+2]))
                            code=""
                            cod=code.join(st)
                            aa = codon_table[cod]
                            ls.append("p.{}{}fs".format(aa,int(int(l[0])/3)+1))
                            Domain(1)#R餘數進位
                            ls.append('/t')
                            ls.append('Frameshift')
                            ls.append('Small deletion')
                            print(ls)
                            csvWrite.writerow([ls])
                        else:
                            st=(str(seq[(int(l[0])-1)]),str(seq[(int(l[0])-1)+1]),str(seq[(int(l[0])-1)+2]))
                            st2=(str(seq[(int(l[1])-1)]),str(seq[(int(l[1])-1)+1]),str(seq[(int(l[1])-1)+2]))
                            code=""
                            code2=""
                            cod=code.join(st)
                            cod2=code2.join(st2)
                            aa = codon_table[cod]
                            aa2 = codon_table[cod2]
                            ls.append("p.{}{}_{}{}del".format(aa,int(int(l[0])/3)+1,aa2,int(int(l[1])/3)+1))
                            for j in listreport2:
                                if int(j[1])<=int(int((int(l[0])+int(l[1]))/2/3))<=int(j[2]):
                                    ls.append(j[0])
                                else:
                                    continue
                            ls.append('/t')
                            ls.append('InFrame deletion')
                            ls.append('Gross deletion')
                            print(ls)
                            csvWrite.writerow([ls])
                            
                    else:
                        st=(str(seq[(int(l[0])-1)]),str(seq[(int(l[0])-1)+1]),str(seq[(int(l[0])-1)+2]))
                        code=""
                        cod=code.join(st)
                        aa = codon_table[cod]
                        ls.append("p.{}{}fs".format(aa,int(int(l[0])/3)+1))
                        Domain(1)#R餘數進位
                        ls.append('/t')
                        ls.append('Frameshift')
                        ls.append('Small deletion')
                        print(ls)
                        csvWrite.writerow([ls])
                    
                elif int(l[0])% 3 == 2:
                    if "_" in list(cell.value):
                        if int(int(l[1])-int(l[0]))<=10:
                            st=(str(seq[(int(l[0])-1)-1]),str(seq[(int(l[0])-1)]),str(seq[(int(l[0])-1)+1]))
                            code=""
                            cod=code.join(st)
                            aa = codon_table[cod]
                            ls.append("p.{}{}fs".format(aa,int(int(l[0])/3)+1))
                            Domain(1)#R餘數進位
                            ls.append('/t')
                            ls.append('Frameshift')
                            ls.append('Small deletion')
                            print(ls)
                            csvWrite.writerow([ls])
                        else:
                            st=(str(seq[(int(l[0])-1)-1]),str(seq[(int(l[0])-1)]),str(seq[(int(l[0])-1)+1]))
                            st2=(str(seq[(int(l[1])-1)-1]),str(seq[(int(l[1])-1)]),str(seq[(int(l[1])-1)+1]))
                            code=""
                            code2=""
                            cod=code.join(st)
                            cod2=code2.join(st2)
                            aa = codon_table[cod]
                            aa2 = codon_table[cod2]
                            ls.append("p.{}{}_{}{}del".format(aa,int(int(l[0])/3)+1,aa2,int(int(l[1])/3)+1))
                            for j in listreport2:
                                if int(j[1])<=int(int((int(l[0])+int(l[1]))/2/3))<=int(j[2]):
                                    ls.append(j[0])
                                else:
                                    continue
                            ls.append('/t')
                            ls.append('InFrame deletion')
                            ls.append('Gross deletion')
                            print(ls)
                            csvWrite.writerow([ls])
                            
                    else:
                        st=(str(seq[(int(l[0])-1)-1]),str(seq[(int(l[0])-1)]),str(seq[(int(l[0])-1)+1]))
                        code=""
                        cod=code.join(st)
                        aa = codon_table[cod]
                        ls.append("p.{}{}fs".format(aa,int(int(l[0])/3)+1))
                        Domain(1)#R餘數進位
                        ls.append('/t')
                        ls.append('Frameshift')
                        ls.append('Small deletion')
                        print(ls)
                        csvWrite.writerow([ls])
                    
                else :
                    continue

        elif "s" in list(cell.value):
            Exon()#def Exon()
            ls.append(cell.value)
            
            if int(l[0]) % 3 == 0:
                st=(str(seq[(int(l[0])-1)-2]),str(seq[(int(l[0])-1)-1]),str(seq[int(l[0])-1]))
                code=""
                cod=code.join(st)
                aa = codon_table[cod]
                ls.append("p.{}{}fs".format(aa,int(int(l[0])/3)))
                Domain(0)#R餘數進位
                ls.append('/t')
                ls.append('Frameshift')
                ls.append('Small insertion')
                print(ls)
                csvWrite.writerow([ls])
            elif int(l[0]) % 3 == 1:
                st=(str(seq[(int(l[0])-1)]),str(seq[(int(l[0])-1)+1]),str(seq[(int(l[0])-1)+2]))
                code=""
                cod=code.join(st)
                aa = codon_table[cod]
                ls.append("p.{}{}fs".format(aa,int(int(l[0])/3)+1))
                Domain(1)#R餘數進位
                ls.append('/t')
                ls.append('Small insertion')
                ls.append('Frameshift')
                print(ls)
                csvWrite.writerow([ls])
            elif int(l[0]) % 3 == 2:
                st=(str(seq[(int(l[0])-1)-1]),str(seq[(int(l[0])-1)]),str(seq[(int(l[0])-1)+1]))
                code=""
                cod=code.join(st)
                aa = codon_table[cod]
                ls.append("p.{}{}fs".format(aa,int(int(l[0])/3)+1))
                Domain(1)#R餘數進位
                ls.append('/t')
                ls.append('Frameshift')
                ls.append('Small insertion')
                print(ls)
                csvWrite.writerow([ls])

        elif "+" in list(cell.value):
            Intronplus()#c.G4816+1A#def Intron()
            ls.append("c.{}+{}{}>{}".format(int(l[0]),int(l[1]),mute[1],mute[2]))   
            if int(l[1])<=2:
                if int(l[0])% 3 == 0:
                    ls.append('/t')
                    Domain(0)#R餘數進位
                    ls.append('/t')
                    ls.append('Splice')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
                elif int(l[0])% 3 == 1:
                    ls.append('/t')
                    Domain(1)#R餘數進位
                    ls.append('/t')
                    ls.append('Splice')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
                elif int(l[0])% 3 == 2:
                    ls.append('/t')
                    Domain(1)#R餘數進位
                    ls.append('/t')
                    ls.append('Splice')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
                else :
                    continue
            elif 2<int(l[1])<=10:
                if int(l[0])% 3 == 0:
                    ls.append('/t')
                    Domain(0)#R餘數進位
                    ls.append('/t')
                    ls.append('Splice region')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
                elif int(l[0])% 3 == 1:
                    ls.append('/t')
                    Domain(1)#R餘數進位
                    ls.append('/t')
                    ls.append('Splice region')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
                elif int(l[0])% 3 == 2:
                    ls.append('/t')
                    Domain(1)#R餘數進位
                    ls.append('/t')
                    ls.append('Splice region')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
            elif 10<int(l[1]):
                if int(l[0])% 3 == 0:
                    ls.append('/t')
                    Domain(0)#R餘數進位
                    ls.append('/t')
                    ls.append('Substitution')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
                elif int(l[0])% 3 == 1:
                    ls.append('/t')
                    Domain(1)#R餘數進位
                    ls.append('/t')
                    ls.append('Substitution')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
                elif int(l[0])% 3 == 2:
                    ls.append('/t')
                    Domain(1)#R餘數進位
                    ls.append('/t')
                    ls.append('Substitution')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
            else:
                print('error')

        elif "-" in list(cell.value):
            Intronminus()#c.G4816-1A#def Intron()
            ls.append("c.{}-{}{}>{}".format(int(l[0]),int(l[1]),mute[1],mute[2])) 
            ls.append('/t')
            if int(l[1])<=2:
                if int(l[0])% 3 == 0:
                    Domain(0)#R餘數進位
                    ls.append('/t')
                    ls.append('Splice')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
                elif int(l[0])% 3 == 1:
                    Domain(1)#R餘數進位
                    ls.append('/t')
                    ls.append('Splice')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
                elif int(l[0])% 3 == 2:
                    Domain(1)#R餘數進位
                    ls.append('/t')
                    ls.append('Splice')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
                else :
                    continue
            elif 2<int(l[1])<=10:
                if int(l[0])% 3 == 0:
                    Domain(0)#R餘數進位
                    ls.append('/t')
                    ls.append('Splice region')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
                elif int(l[0])% 3 == 1:
                    Domain(1)#R餘數進位
                    ls.append('/t')
                    ls.append('Splice region')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
                elif int(l[0])% 3 == 2:
                    Domain(1)#R餘數進位
                    ls.append('/t')
                    ls.append('Splice region')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
            elif 10<int(l[1]):
                if int(l[0])% 3 == 0:
                    Domain(0)#R餘數進位
                    ls.append('/t')
                    ls.append('Substitution')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
                elif int(l[0])% 3 == 1:
                    Domain(1)#R餘數進位
                    ls.append('/t')        
                    ls.append('Substitution')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
                elif int(l[0])% 3 == 2:
                    Domain(1)#R餘數進位
                    ls.append('/t')        
                    ls.append('Substitution')
                    ls.append('Transcription')
                    print(ls)
                    csvWrite.writerow([ls])
                    
                    
        else: 
            mute = ''.join(re.findall(r'[A-Za-z]', cell.value))#取字母
            pospy=int(l[0])-1
            #csvWrite.writerow(["NM_000138.4:c.{}{}>{}".format(int(l[0]),seq[pospy],mute[2])])
            Exon()#def Exon()
            ls.append("c.{}{}>{}".format(int(l[0]),seq[pospy],mute[2]))    
                
            if int(l[0]) % 3 == 0:
                st=(str(seq[pospy-2]),str(seq[pospy-1]),str(seq[pospy]))
                stm=(str(seq[pospy-2]),str(seq[pospy-1]),str(mute[2]))
                code=""
                codem=""
                cod=code.join(st)
                codm=codem.join(stm)
                aa = codon_table[cod]
                aam = codon_table[codm]
                ls.append("p.{}{}{}".format(aa,int(int(l[0])/3),aam))
                if aa == aam:
                    Domain(0)#R餘數進位
                    ls.append('/t')        
                    ls.append("Silence")
                    ls.append("Substution")
                    csvWrite.writerow([ls])
                    print(ls)
                else :
                    if aam == "*":
                        Domain(0)#R餘數進位
                        ls.append('/t')
                        ls.append("Nonsense")
                        ls.append("Substution")
                        csvWrite.writerow([ls])
                        print(ls)
                    else:
                        for j in listreport2:
                            if int(j[1])<=int(int(l[0])/3)<=int(j[2]):
                                ls.append(j[0])
                                if (j[0][0]) == "c":
                                    Count=1
                                    for f in listFeatures:
                                        if int(int(l[0])/3)==int(f):
                                            ls.append("1")
                                            break
                                        elif Count==516:
                                            ls.append("0")
                                        else:
                                            Count+=1
                                            continue
                                else:
                                    ls.append('/t')
                            else:
                                continue       
                        ls.append("Missense")
                        ls.append("Substution")
                        csvWrite.writerow([ls])
                        print(ls)
                            
            elif int(l[0]) % 3 == 1:
                st=(str(seq[pospy]),str(seq[pospy+1]),str(seq[pospy+2]))
                stm=(str(mute[2]),str(seq[pospy+1]),str(seq[pospy+2]))
                code=""
                codem=""
                cod=code.join(st)
                codm=codem.join(stm)
                aa = codon_table[cod]
                aam = codon_table[codm]
                if mute[2] == "":
                    print("Frameshift")
                else:
                    ls.append("p.{}{}{}".format(aa,int(int(l[0])/3) + 1,aam))
                    if aa == aam:
                        Domain(1)#R餘數進位
                        ls.append('/t')
                        ls.append("Silence")
                        ls.append("Substution")
                        csvWrite.writerow([ls])
                        print(ls)
                    else :
                        if aam == "*":
                            Domain(1)#R餘數進位
                            ls.append('/t')
                            ls.append("Nonsense")
                            ls.append("Substution")
                            csvWrite.writerow([ls])
                            print(ls)
                        else:
                            for j in listreport2:
                                if int(j[1])<=int(int(l[0])/3)+1<=int(j[2]):
                                    ls.append(j[0])
                                    if (j[0][0]) == "c":
                                        Count=1
                                        for f in listFeatures:
                                            if int(int(l[0])/3)+1==int(f):
                                                ls.append("1")
                                                break
                                            elif Count==516:
                                                ls.append("0")
                                            else:
                                                Count+=1
                                                continue
                                    else:
                                        ls.append('/t')
                                else:
                                     continue
                            ls.append("Missense")
                            ls.append("Substution")
                            csvWrite.writerow([ls])
                            print(ls)
                            
            elif int(l[0]) % 3 == 2:
                stm=(str(seq[pospy-1]),str(mute[2]),str(seq[pospy+1]))
                st=(str(seq[pospy-1]),str(seq[pospy]),str(seq[pospy+1]))
                code=""
                codem=""
                cod=code.join(st)
                codm=codem.join(stm)
                aa = codon_table[cod]
                aam = codon_table[codm]
                if mute[2] == "":
                    print("Frameshift")
                else:
                    ls.append("p.{}{}{}".format(aa,int(int(l[0])/3) + 1,aam))
                    if aa == aam:
                        Domain(1)#R餘數進位
                        ls.append('/t')
                        ls.append("Silence")
                        ls.append("Substution")
                        csvWrite.writerow([ls])
                        print(ls)
                    else :
                        if aam == "*":
                            Domain(1)#R餘數進位
                            ls.append('/t')
                            ls.append("Nonsense")
                            ls.append("Substution")
                            csvWrite.writerow([ls])
                            print(ls)
                        else:
                            for j in listreport2:
                                if int(j[1])<=int(int(l[0])/3)+1<=int(j[2]):
                                    ls.append(j[0])
                                    if (j[0][0]) == "c":
                                        Count=1
                                        for f in listFeatures:
                                            if int(int(l[0])/3)+1==int(f):
                                                ls.append("1")
                                                break
                                            elif Count==516:
                                                ls.append("0")
                                            else:
                                                Count+=1
                                                continue
                                    else:
                                        ls.append('/t')
                                else:
                                    continue
                            ls.append("Missense")
                            ls.append("Substution")
                            csvWrite.writerow([ls])
                            print(ls)
            else:
                print("error")
fn.close()
fncsv2.close()
fncsv1.close()
csvfn.close()

