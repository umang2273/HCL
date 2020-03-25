# -*- coding: utf-8 -*-
"""
Created on Tue Mar 24 10:51:38 2020

@author: UJP
"""
import nltk
import xlwt
import os
files = [f for f in os.listdir('.') if os.path.isfile(f)]
num = []
for i in range(10):
    num.append(str(i))
num = num + ['-']
book = xlwt.Workbook()   
sheet = book.add_sheet('sheet 1') 
#row = 0
sheet.write(0, 0, 'Filename')   
sheet.write(0, 1, 'Extracted Values') 
row = 1
format = '.txt'
#stopping = ['STATEMENTS','Director','For']
for f in files:
    try:
        if (format == f[-4:]):
            f1=open(f)  
            lines = f1.readlines()
            tk_1 = nltk.word_tokenize(lines[1])
    #print(tk_1)
            d = {}
            #print(f)
            tk_n = []
            for i in range(len(tk_1)):
                if tk_1[i][0] in num:
                    tk_n.append(tk_1[i])
            #print(tk_n)
            for i in range(len(tk_n)):
                if tk_n[i] == '2019':
                    p = i
                    break;
                else:
                    p = -1
            for j in range(3,len(lines)):
                tk = nltk.word_tokenize(lines[j])
                text = ''
                value = []
                for i in range(len(tk)):
                    if tk[i][0] in num:
                        tk[i].replace(',','')
                        value.append(tk[i])
                    elif tk[i][0] == '(':
                        if tk[i][1] in num:
                            value.append(tk[i])
                        else:
                            text = text+tk[i]+" "
                    else:
                        text = text+tk[i]+" "
                text = text[:-1]
                # for i in range(len(value)):
                #     if value[i][0] == '(':
                #         value[i][0] = '-'
                #         del(value[i][-1])
                
                #print(value)
                if p == -1:
                    if value != []:
                        d[text] = 'nan'
                else:
                    if value != []:
                        d[text] = value[p]
                if 'STATEMENTS' in tk:
                    break;
                elif 'Director' in tk:
                    break;
                elif 'For' in tk:
                    break;
            fname = f[:-4]   
            sheet.write(row, 0, fname)   
            sheet.write(row, 1, str(d))   
            # incrementing the value of row by one with each iterations.   
            row = row + 1  
            
            #book.close()
    # if year in lines[1]:
    #     count+=1
            f1.close()
            book.save('Results.csv')
    except (IndexError):
        fname = f[:-4]  
            
        sheet.write(row, 0, fname)   
        sheet.write(row, 1, str(d))   
        # incrementing the value of row by one with each iterations.   
        row = row + 1  
        #book.close()
# if year in lines[1]:
#     count+=1
        f1.close()
        book.save('Results.csv')
            
                
        # Rows and columns are zero indexed.  
       
        
# print(count)

# for line in lines:
#     print(line.strip())
# f.close()
# for i in range(len(data)):
#     print (data[i])

# nltk_tokens = nltk.word_tokenize(data)
# print (nltk_tokens)
 

  

