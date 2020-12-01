import pandas as pd
import datetime
import smtplib
import os
os.chdir(r"D:\Automated Birthday Wisher")
# os.mkdir("anans")


GMAIL_ID = 'anandrawool9999@gmail.com'
GMAIL_PSWD = 'Anand@2807'


def sendEmail(to,sub,msg):
    print(f"Email to {to} send with subject: {sub} and message {msg}")
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PSWD)
    s.sendmail(GMAIL_ID,to, f"Subject:{sub}\n\n{msg}")
    s.quit()
    

if __name__ == "__main__":
    # sendEmail('anandrawool9999@gmail.com',"HAPPY BIRTHDAY", "test Message")
    # exit()
    dr=pd.read_excel("data.xlsx")
    # print(dr)
    today=datetime.datetime.now().strftime("%d-%m")
    yearNow=datetime.datetime.now().strftime("%Y")
    # print(today)
    writeInd = []

    for index, item in dr.iterrows():
        # print(index,item['Birthday'])
        bday=item['Birthday'].strftime("%d-%m")
        # print(bday)
        if(today==bday and yearNow not in str(item['Year'])):
            sendEmail(item['Email'],"Happy Birthday",item['Dailogue'])
            writeInd.append(index)
    # print(writeInd)

    for i in writeInd:
        yr = dr.loc[i,'Year']
        dr.loc[i,'Year']= str(yr) + ',' + str(yearNow)
        # print(yr)
    # print(dr)
    dr.to_excel('data.xlsx',index=False)


    

            

    





