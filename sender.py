import sys
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import pandas as pd
import logging
from datetime import datetime

def genderToTitle(gender:str):
    if gender == "Female":
        return "Ms."
    elif gender == "Male":
        return "Mr."
    else:
        return None

def fixSurname(surname:str):
    return surname[0].upper() + surname[1:].lower()

def main():
    try:
        with open("check_today.txt", "r") as f:
            try:
                if f.readlines()[-1] == datetime.today().strftime('%Y-%m-%d'):
                    logger.exception("Today's emails are already sent.")
                    sys.exit("Today's emails are already sent.")
            except IndexError:
                pass
    except FileNotFoundError:
        tempp = open("check_today.txt", "w")
        tempp.close()
        logger.info("check_today.txt created.")
        print("\ncheck_today.txt doesn't exist, a new one is created.")
        print("Do you want to continue? yes/no")
        while 1:
            condition = input()
            if condition == "y" or condition == "yes":
                break
            elif condition == "n" or condition == "no":
                sys.exit("Exiting...")
            print("Try again. Do you want to continue? yes/no\n")
        
    smtp_ssl_host = 'smtp.metu.edu.tr'
    smtp_ssl_port = 465
    username = 'eXXXXXX'
    password = '********'
    sender = 'Abdullah Emir Gogusdere <eXXXXXX@metu.edu.tr>'
    save_mail = 'abemirgo@gmail.com'

    logger.info(", ".join([smtp_ssl_host, smtp_ssl_host, username, sender, save_mail]))

    chosens = []
    profList = pd.read_csv('deneme.csv')
    logger.info(profList)

    notSentprof = profList[profList.sent == 0]

    if len(notSentprof) == 0:
        logger.log(logging.ERROR, 'There are no email addresses in the csv file that have not been sent.', exc_info=True)
        sys.exit("There are no email addresses in the csv file that have not been sent.")
    
    univs = set(notSentprof["university"].to_list())
    for uni in univs:
        temp = notSentprof[notSentprof.university == uni].iloc[[0]]
        temp_index = temp.index[0]
        chosens.append([temp.iloc[0], temp_index])

    with open("check_today.txt", "a") as f:
        f.write("\n")
        f.write(datetime.today().strftime('%Y-%m-%d'))  
    logger.info(f"Today's date ({datetime.today().strftime('%Y-%m-%d')}) is entered to check_today.txt")
    
    for lst in chosens:
        chosenOne, index = lst[0], lst[1] 
        target = chosenOne["email"]
        titleOfRespect = genderToTitle(chosenOne["gender"])
        SurnameOfProf = fixSurname(chosenOne["surname"])

        to = target
        bcc = [save_mail]

        target = [to] + bcc
        titleOfRespect = genderToTitle(chosenOne["gender"])
        SurnameOfProf = fixSurname(chosenOne["surname"])

        msg = MIMEMultipart()
        msg['Subject'] = 'Summer Internship'
        logger.info('Subject: ' + msg['Subject'])
        msg['From'] = sender
        logger.info('From: ' + msg['From'])
        msg['To'] = to
        logger.info('To: ' + msg['To'])

        with open("message.txt", "r", encoding="utf-8") as f:
            if titleOfRespect:
                plaintxt = f"Dear {titleOfRespect} {SurnameOfProf},\n\n" + f.read()
            else:
                plaintxt = f"Dear {SurnameOfProf},\n\n" + f.read()

        logger.info("Message:\n" + plaintxt)
        txt = MIMEText(plaintxt)
        msg.attach(txt)

        file_path = "resumeAbdullah.pdf"
        file_name = "resumeAbdullah.pdf"

        with open(file_path, "rb") as file:
            part = MIMEApplication(file.read())
        part['Content-Disposition'] = f'attachment; filename="{file_name}"'
        msg.attach(part)
        logger.info("Attached file path/name: {}/{}".format(file_path, file_name))

        try:
            server = smtplib.SMTP_SSL(smtp_ssl_host, smtp_ssl_port)
            server.login(username, password)
            print(f"Connected to {smtp_ssl_host}:{smtp_ssl_port}")
            logger.info(f"Connected to {smtp_ssl_host}:{smtp_ssl_port}")
        except Exception:
            logger.exception(f"Couldn't connect to {smtp_ssl_host}:{smtp_ssl_port} with the username {username} or the password")
        
        try:
            server.sendmail(sender, target, msg.as_string())
            print(f"E-mail sent to {target}.")
            logger.info(f"E-mail sent to {target}.")
        except Exception:
            logger.exception("Couldn't send the e-mail.")

        try:
            profList.at[index, 'sent'] = 1
            logger.info(profList)
            profList.to_csv("deneme.csv", index=False)
        except Exception:
            logger.exception("Couldn't save the csv file.")

    server.quit()
    logging.shutdown()

if __name__ == "__main__":
    logging.basicConfig(filename="email.log",
                            filemode='a',
                            format='%(asctime)s:%(msecs)d %(name)s %(levelname)s %(message)s',
                            datefmt='"%Y-%m-%d %H:%M:%S',
                            level=logging.DEBUG)
    logging.info("Running Auto Email Sender")

    logger = logging.getLogger('emailWatch')
    main()
