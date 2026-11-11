"""Sample 10:10 IST production reminder dispatcher.
Safe default: preview only. Requires private self-hosted runner with workbook access.
Dependencies: openpyxl, requests. No third-party mail service.
"""
import datetime as dt, json, os, pathlib, sys
from zoneinfo import ZoneInfo
import openpyxl
import requests

RECIPIENTS=["finny.jason@beeteam.in","jawahar.muthu@beeteam.in"]
IST=ZoneInfo("Asia/Kolkata")

def items(path, now=None):
    today=(now or dt.datetime.now(IST)).date()
    book=openpyxl.load_workbook(path,read_only=True,data_only=True)
    sheet=book["Process Workflow"]
    results=[]
    for row in list(range(11,31))+list(range(36,41)):
        name=sheet.cell(row,1).value
        when=sheet.cell(row,2).value
        status=str(sheet.cell(row,6).value or "").strip()
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
    tenant=os.environ["MS_TENANT_ID"]
    client=os.environ["MS_CLIENT_ID"]
    secret=os.environ["MS_CLIENT_SECRET"]
    sender=os.environ["MAIL_SENDER_UPN"]
    token=requests.post(f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={"grant_type":"client_credentials","client_id":client,"client_secret":secret,
              "scope":"https://graph.microsoft.com/.default"},timeout=25)
    token.raise_for_status()
    mail={"message":{"subject":"Production Tracker: incomplete tasks at 10:10 AM IST",
        "body":{"contentType":"Text","content":body},
        "toRecipients":[{"emailAddress":{"address":x}} for x in RECIPIENTS]},
        "saveToSentItems":True}
    response=requests.post(f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail",
        json=mail,headers={"Authorization":"Bearer "+token.json()["access_token"]},timeout=30)
    response.raise_for_status()

def main():
    path=os.environ.get("WORKBOOK_PATH")
    if not path or not pathlib.Path(path).is_file():
        print("Not configured: WORKBOOK_PATH must point to the LIVE accessible private workbook.")
        return 2
    pending=items(path)
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
    log=pathlib.Path(os.environ["REMINDER_LOG_PATH"])
    log.parent.mkdir(parents=True,exist_ok=True)
    if log.exists() and day in log.read_text(encoding="utf-8").splitlines():
        print("Already sent on this date; skipping.")
        return 0
    deliver(body)
    with log.open("a",encoding="utf-8") as f:f.write(day+"\n")
    print("Email accepted by Microsoft Graph.")
    return 0

if __name__=="__main__":sys.exit(main())
