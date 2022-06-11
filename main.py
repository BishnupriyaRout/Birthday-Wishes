import pandas as pd 
import datetime
import smtplib
import os
os.chdir(r"C:\Python310\Birthday Wishes")

#Enter your authentication details
GMAIL_ID = '' #not giving for security reasons
GMAIL_PSWD = ''  #not giving for security reasons

def sendEmail(to,sub,msg):
    print(f"Send email to {to} sent with subject: {sub} and message: {msg}")
    s = smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PSWD)
    s.sendmail(GMAIL_ID, to , f'Subject: {sub}\n\n{msg}')
    s.quit()

if __name__ == "__main__":
    sendEmail(GMAIL_ID,"subject","test message")
    df = pd.read_excel("data.xlsx")
    # print(df)
    today = datetime.datetime.now().strftime("%d-%m")
    yearNow = datetime.datetime.now().strftime("%Y")
    # print(today)
    writeInd = []

    for index,item in df.iterrows():

        # print(index,item['Birthday'])
        bday = item['Birthday'].strftime("%d-%m")
        # print(bday)
        # msg = "I don't usually wish birthdays , but you are exceptional !!."

        if today == bday and yearNow not in str(item['Year']):
            sendEmail(item['Email'], "HAPPY BIRTHDAY",item['Dialogue'])
            writeInd.append(index)

    if writeInd:
        for i in writeInd:
            yr = df.loc[i,'Year']
            df.loc[i,'Year'] = str(yr) + ',' + str(yearNow)
            # print(df.loc[i,'Year'])

        # print(df)
        df.to_excel('data.xlsx',index=False)