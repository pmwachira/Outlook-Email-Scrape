import pandas as pd
import re
import unicodedata
from bs4 import BeautifulSoup
def scraper(html,message_subject):
     #replace class=MsoNormal with class="MsoNormal"
    #replace <span lang=EN-GB> with <span lang="EN-GB">
    replaced_text=html.replace('MsoNormal','\"MsoNormal\"').replace('EN-GB','\"EN-GB\"').replace(u'\xa0', u' ').replace("\t"," ").rstrip("\r\n").replace('\r',' ')
    #re.sub(r'(\r\n.?)+', r'\r\n', text)
    soup = BeautifulSoup(replaced_text,"lxml")

    count=0
    row_data=[]
    tag_series=soup.findAll('table', width = 0)
    
   
    check=False

    for entry in tag_series:
        #highlight required tags
        if(count==0):
                    count_internal=0
                    split_this=re.sub(r'(\n\n)+', r'\n',unicodedata.normalize("NFKD", entry.text.replace('GSOCSecurity Incident','|*Security Incident|*').replace('Priority:','|*Priority:|*').replace('Incident Title:','|*Incident Title:|*').replace('Incident Description:','|*Incident Description:|*')\
                    .replace('Event Timestamp:','|*Event Timestamp:|*').replace('Detection Method:','|*Detection Method:|*').replace('Threat Category:','|*Threat Category:|*').replace('Remediation Actions','|*Remediation Actions|*')\
                    .replace('Incident Details','|*Incident Details|*')))
                    splitted_array=split_this.split('|*')
                    aSize=len(splitted_array)-1
                    
                    for x in splitted_array:
                        #handle mail template 2
                        if x=='Priority:'  and  count_internal==1:
                        
                            row_data.insert(0,message_subject.split('|')[0].upper().replace(' ','').replace(':','').replace('RE','').replace('SECURITYINCIDENT',''))
                        if x=='Incident Title:'  and  count_internal==2:
                            row_data.insert(1,splitted_array[count_internal+1])
                        if x=='Incident Description:'  and  count_internal==3:
                            row_data.insert(2,splitted_array[count_internal+2])

                       
                  

                                    
                           
                    
                                
                        else:
                            #handle mail template 1
                            if(x=='Security Incident'):
                                
                                if(count>=aSize):
                                    row_data.insert(0,'')
                                else:
                                    row_data.insert(0,splitted_array[count_internal+1])
                            if(x=='Priority:'):
                                if(count>=aSize):
                                    row_data.insert(1,'')
                                else:
                                    row_data.insert(1,splitted_array[count_internal+1])
                            if(x=='Incident Title:'):
                                if(count>=aSize):
                                    row_data.insert(2,'')
                                else:
                                    row_data.insert(2,splitted_array[count_internal+1])
                            if(x=='Incident Description:'):
                                if(count>=aSize):
                                    row_data.insert(3,'')
                                else:
                                    row_data.insert(3,splitted_array[count_internal+1])
                            if(x=='Event Timestamp:'):
                                if(count>=aSize):
                                    row_data.insert(4,'')
                                else:
                                    row_data.insert(4,splitted_array[count_internal+1])
                            if(x=='Detection Method:'):
                                if(count>=aSize):
                                    row_data.insert(5,'')
                                else:
                                    row_data.insert(5,splitted_array[count_internal+1])
                            if(x=='Threat Category:'):
                                if(count>=aSize):
                                    row_data.insert(6,'')
                                else:
                                    row_data.insert(6,splitted_array[count_internal+1])
                            if(x=='Remediation Actions'):
                                if(count>=aSize):
                                    row_data.insert(7,'')
                                else:
                                    row_data.insert(7,splitted_array[count_internal+1])
                            if(x=='Incident Details'):
                                if(count>=aSize):
                                    row_data.insert(8,'')
                                else:
                                    row_data.insert(8,splitted_array[count_internal+1])

                            count_internal=count_internal+1        

        count=count+1

                       
    if(len(row_data)>9):                          
        print(row_data)
    return row_data
    

import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)


test = inbox.Folders(14).Folders(1)

messages = test.Items
count=0


all_row_data=[]
countx=0
##for testmessageindex in range(10):
for testmessageindex in range(messages.Count):
    done=False
    message_subject=messages[testmessageindex].Subject
    for attachmentinstance in range(messages[testmessageindex].Attachments.Count):
        
        if(messages[testmessageindex].Attachments[attachmentinstance].FileName).upper().endswith('.PDF'):
            
            message = messages[testmessageindex]
            body_content = message.HTMLBody         
            row_data=scraper(body_content,message_subject)
            all_row_data.append(row_data)

    
            done=True
            countx=countx+1
            
        if done:break
                    
column_headers=['Security Incident','Priority','Incident Title','Incident Description','EventTimestamp','Detection Method','Threat Category','Remediation Actions','Incidence Details']


print('countx -> '+str(countx))

df = pd.DataFrame(all_row_data, columns=column_headers)

df.to_excel("test_table.xlsx")
