# -- coding: utf-8 --


def to_timestamp(date_time):
    import time
    timearray = time.strptime(date_time, "%Y-%m-%d %H:%M:%S")
    timestamp = int(time.mktime(timearray))
    return str(timestamp) + "000"


def download_grafana_img(from_ts, to_ts, panelid, file_name):
    import requests
    grafana_url = "http://ip:23000"
    grafana_path = "{0}/render/d-solo/XXXX/xun-jian-da-pan".format(grafana_url)
    grafana_args = "?from={0}&to={1}&orgId=1&panelId={2}&width=1000&height=500&tz=Asia%2FShanghai".format(from_ts,
                                                                                                          to_ts,
                                                                                                          panelid)
    img_url = grafana_path + grafana_args
    print(img_url)
    r = requests.get(img_url, stream=True)
    if r.status_code == 200:
        open(file_name, 'wb').write(r.content)


def obs_upload(objectfile, uploadfile):
    from obs import ObsClient

    AK = "xxx"
    SK = "xxx"
    ENDPOINT = "obs.cn-south-1.myhuaweicloud.com"

    obsClient = ObsClient(access_key_id=AK, secret_access_key=SK, server=ENDPOINT)

    bucket_name = "grafana-pics"
    remote_prefix = "grafana_reports"

    objectkey = remote_prefix + objectfile
    obsClient.uploadFile(bucket_name, objectkey, uploadfile, encoding_type="url")


def path_to_image_html(path):
    return '<img src="' + path + '" width="200" />'


def send_mail(html):
    import smtplib
    from email.mime.text import MIMEText
    from email.header import Header

    mail_host = "smtp.test.com.cn"
    mail_user = "itservice@test.com.cn"
    mail_pass = "xxx"

    sender = 'itservice@test.com.cn'
    receivers = ['xxx@test.com.cn']

    message = MIMEText(html, 'html', 'utf-8')
    message['From'] = sender
    message['To'] = ','.join(receivers)

    subject = '巡检报告'
    message['Subject'] = Header(subject, 'utf-8')

    try:
        smtpObj = smtplib.SMTP()
        smtpObj.connect(mail_host, 25)
        smtpObj.login(mail_user, mail_pass)
        smtpObj.sendmail(sender, receivers, message.as_string())
        print("邮件发送成功")
    except smtplib.SMTPException as e:
        print(e)


if __name__ == '__main__':
    import os
    import datetime
    import pandas as pd

    date_cur = datetime.datetime.now().strftime("%Y%m%d")
    date_now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    date_tmp = datetime.datetime.now() - datetime.timedelta(days=1)
    date_ago = date_tmp.strftime("%Y-%m-%d %H:%M:%S")

    date_now_ts = to_timestamp(date_now)
    date_ago_ts = to_timestamp(date_ago)

    df = pd.read_excel('./xlsx/巡检模板.xlsx')

    img_path_list = []
    img_path = "./img/" + date_cur
    if not os.path.exists(img_path):
        os.mkdir(img_path)
    for img_id in df['图表ID']:
        file_name = img_path + "/" + str(img_id) + ".png"
        download_grafana_img(date_ago_ts, date_now_ts, img_id, file_name)
        objectfile = "/" + date_cur + "/" + str(img_id) + ".png"
        obs_url = "https://grafana.obs.cn-south-1.myhuaweicloud.com/grafana_reports" + objectfile
        img_path_list.append(obs_url)
        obs_upload(objectfile, file_name)

    df["巡检结果图片"] = img_path_list

    table = df.to_html(escape=False, formatters=dict(巡检结果图片=path_to_image_html))

    html_template = """
    <!doctype html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>巡检报告</title>
        <style type="text/css">
            .table_css table{{
                border-collapse: collapse;
                width:100%;
                border:1px solid #B8B8B8 !important;
                margin-bottom:20px;
            }}
            .table_css th{{
                background-color:#DDEEFF !important;
                padding:5px 9px;
                font-size:14px;
                font-weight:normal;
                text-align:center;
            }}
            .table_css td{{
                padding:5px 9px;
                font-size:12px;
                font-weight:normal;
                text-align:center;
                word-break: break-all;
            }}
            .table_css tr:nth-child(odd){{
                background-color:#FFFFFF !important;
            }}
            .table_css tr:nth-child(even){{
                background-color: #E6E6E6 !important;
            }}
        </style>
    </head>
    <body>
    <h3  class="textCetent">巡检报告</h3>
    <div class="table_css">
        {table}
    </div>
    </body>
    </html>
    """

    html = html_template.format(table=table)

    with open("./html/巡检_{0}.html".format(date_cur), "w", encoding='utf-8') as f:
        f.write(html)

    send_mail(html)
