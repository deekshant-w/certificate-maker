def createAndSend(send_to,certiFile,send_from,subject,message,password):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Subject'] = subject
    msg.attach(MIMEText(message))
    part = MIMEBase('application', "octet-stream")
    with open(certiFile, 'rb') as file:
        part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition','attachment; filename="{}"'.format(Path(certiFile).name))
    msg.attach(part)
    smtp = smtplib.SMTP(host="smtp.gmail.com", port= 587)
    smtp.starttls()
    smtp.login(send_from, password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()


def email(thisButton,controller):
    wb = oxl.load_workbook(fnf['projectDatabase'])
    ws = wb.active
    data = getData(ws)
    if('EMAIL_ID' in data[0]):
        thisButton.configure(background='SystemButtonFace')
        thisButton.configure(text='Email Certificates')
        pdfs = [f for f in os.listdir(fnf['projectCertificates']) if f.split(".")[-1]=='pdf']
        emailIndex = data[0].index('EMAIL_ID')
        ids = []
        for x in range(1,len(data)):
            ids.append(data[x][emailIndex])
        counter = 1
        thisButton.configure(text='Please Wait ..')
        eob.configure(text=f"0/{len(ids)}")
        controller.update()
        if(not os.path.exists(fnf['projectSettings'])):
            createSettingsFile(fnf['projectSettings'])
        settings = json.loads(open(fnf['projectSettings']).read())
        send_from = settings['emailFrom']
        subject = settings['emailSubject']
        message = settings['emailBody']
        password = "ldqsjwdxgkzqimsv" #settings['emailPassword']
        for x in range(len(ids)):
            pdfLocation = fnf['projectCertificates']+"/"+pdfs[x]
            createAndSend(ids[x],pdfLocation,send_from,subject,message,password)
            eob.configure(text=f"{counter}/{len(ids)}")
            controller.update()
            counter+=1
        thisButton.configure(text='Email Certificates '+u'\u2713')
        thisButton.configure(bg = "#00ff11")
        eob.configure(text=f"{counter-1}/{len(ids)}")
    else:
        messagebox.showinfo("No Emails Found!","PLease fill the EMAIL_ID field in the database.")


def main_settings(thisButton):
    from sqlalchemy import create_engine
    from sqlalchemy.ext.declarative import declarative_base
    from sqlalchemy import Column, Integer, String, DateTime
    from sqlalchemy.orm import sessionmaker
    from sqlalchemy.sql import func
    import datetime
    if(os.path.exists("./data.sqlite3")):
        inp = messagebox.askyesno("Already Exists","Database already exists! Do you wish to clean all data from the Database?")
        if(inp):
            os.remove(".\\data.sqlite3")
        else:
            return
    engine = create_engine("sqlite:///data.sqlite3",echo=True)
    Base.metadata.create_all(engine)
    sessionmaker(bind=engine)().commit()

    Base = declarative_base()
    class Records(Base):
        __tablename__ = 'records'
        id = Column(Integer, primary_key=True)
        data = Column(String)
        project = Column(String)
        created_date = Column(DateTime(timezone=True), default=datetime.datetime.now)
        def __repr__(self):
            return "<Records(id='%s', data='%s', project='%s', created_date='%s')>" % (self.id, self.data, self.project, self.created_date)
    # engine = create_engine("sqlite:///data.sqlite3",echo=True)
    # ed_user = Records(data='abc')
    # Session = sessionmaker(bind=engine)
    # session = Session()
    # our_user = session.query(Records).filter_by(data='abc').first()
    # print(our_user.created_date)
    # session.add(ed_user)
    # session.commit()
