"""Sample 10:10 IST production reminder dispatcher.
Safe default: preview only. Requires private self-hosted runner with workbook access.
Dependencies: openpyxl, requests. No third-party mail service.
"""
import base64, datetime as dt, io, json, os, pathlib, sys
import openpyxl
import requests

RECIPIENTS=["finny.jason@beeteam.in","jawahar.muthu@beeteam.in"]
IST=dt.timezone(dt.timedelta(hours=5,minutes=30))

def token():
    tenant=os.environ['MS_TENANT_ID']
    response=requests.post(f'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token',data={'grant_type':'client_credentials','client_id':os.environ['MS_CLIENT_ID'],'client_secret':os.environ['MS_CLIENT_SECRET'],'scope':'https://graph.microsoft.com/.default'},timeout=25)
    response.raise_for_status()
    return response.json()['access_token']

def sharepoint_workbook():
    url=os.environ['WORKBOOK_SHARE_URL']
    share_id='u!'+base64.urlsafe_b64encode(url.encode()).decode().rstrip('=')
    response=requests.get(f'https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem/content',headers={'Authorization':'Bearer '+token()},timeout=60)
    response.raise_for_status()
    return io.BytesIO(response.content)

def items(path, now=None):
    today=(now or dt.datetime.now(IST)).date()
    book=openpyxl.load_workbook(path,read_only=True,data_only=True)
    sheet=book["Process Workflow"]
    results=[]
    for row, date_col, status_col in [(r,2,6) for r in range(11,31)] + [(r,2,6) for r in range(36,41)] + [(r,3,4) for r in range(44,51)]:
        name=sheet.cell(row,1).value
        when=sheet.cell(row,date_col).value
        status=str(sheet.cell(row,status_col).value or "").strip()
        if not name or when is None or status.casefold()=="completed":
            continue
        if isinstance(when,dt.datetime): when=when.date()
        if isinstance(when,str):
            try:when=dt.date.fromisoformat(when[:10])
            except ValueError:continue
        if not isinstance(when,dt.date) or when>today:continue
        results.append({"row":row,"task":str(name),"date":when.isoformat(),"status":status or "Not marked"})
    return results

def deliver(body):
    sender=os.environ["MAIL_SENDER_UPN"]
    mail={"message":{"subject":"Production Tracker: incomplete tasks at 10:10 AM IST",
        "body":{"contentType":"Text","content":body},
        "toRecipients":[{"emailAddress":{"address":x}} for x in RECIPIENTS]},
        "saveToSentItems":True}
    response=requests.post(f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail",
        json=mail,headers={"Authorization":"Bearer "+token()},timeout=30)
    response.raise_for_status()

def main():
    url=os.environ.get('WORKBOOK_SHARE_URL')
    path=os.environ.get('WORKBOOK_PATH')
    if url:
        source=sharepoint_workbook()
    elif path and pathlib.Path(path).is_file():
        source=path
    else:
        print('Configure WORKBOOK_SHARE_URL for SharePoint or WORKBOOK_PATH for private drive.')
        return 2
    pending=items(source)
    if not pending:
        print("No dated outstanding tasks. No email.")
        return 0
    now=dt.datetime.now(IST)
    day=now.date().isoformat()
    body="Production tasks not marked Completed as of 10:10 AM IST:\n\n"+"\n".join(
        f"- {x['task']} | {x['date']} | {x['status']} | Excel row {x['row']}" for x in pending)
    print(json.dumps({"date":day,"count":len(pending),"tasks":pending},indent=2))
    if os.getenv("SEND_ENABLED")!="true":
        print("DRY RUN: no email sent.")
        return 0
    # Use persistent storage on the self-hosted runner, not an ephemeral GitHub workspace.
    log=pathlib.Path(os.environ['REMINDER_LOG_PATH'])
    log.parent.mkdir(parents=True,exist_ok=True)
    if log.exists() and day in log.read_text(encoding="utf-8").splitlines():
        print("Already sent on this date; skipping.")
        return 0
    deliver(body)
    with log.open("a",encoding="utf-8") as f:f.write(day+"\n")
    print("Email accepted by Microsoft Graph.")
    return 0

if __name__=="__main__":sys.exit(main())
