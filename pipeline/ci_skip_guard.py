import sys, os
from datetime import date, timedelta
try:
    import datetime as _dt
    from zoneinfo import ZoneInfo
    today = _dt.datetime.now(ZoneInfo("America/Toronto")).date()
except Exception:
    today = date.today()

def nth_dow(y, m, dow, n):
    d = date(y, m, 1)
    off = (dow - d.weekday()) % 7
    return d + timedelta(days=off + (n - 1) * 7)

def observed(d):
    if d.weekday() == 5: return d + timedelta(days=2)
    if d.weekday() == 6: return d + timedelta(days=1)
    return d

def easter(y):
    a=y%19; b=y//100; c=y%100; d=b//4; e=b%4; f=(b+8)//25; g=(b-f+1)//3
    h=(19*a+b-d-g+15)%30; i=c//4; k=c%4; l=(32+2*e+2*i-h-k)%7
    m=(a+11*h+22*l)//451; mo=(h+l-7*m+114)//31; da=((h+l-7*m+114)%31)+1
    return date(y, mo, da)

y = today.year
victoria = date(y,5,24) - timedelta(days=date(y,5,24).weekday())
hols = {
    "New Year's Day": observed(date(y,1,1)),
    "Family Day": nth_dow(y,2,0,3),
    "Good Friday": easter(y) - timedelta(days=2),
    "Victoria Day": victoria,
    "Canada Day": observed(date(y,7,1)),
    "Civic Holiday": nth_dow(y,8,0,1),
    "Labour Day": nth_dow(y,9,0,1),
    "Thanksgiving Day": nth_dow(y,10,0,2),
    "Christmas Day": observed(date(y,12,25)),
}
xd = date(y,12,25).weekday()
if xd in (5,6): hols["Boxing Day"] = observed(date(y,12,25)) + timedelta(days=1)
elif xd == 4:   hols["Boxing Day"] = date(y,12,26) + timedelta(days=2)
else:           hols["Boxing Day"] = date(y,12,26)

name = None
for n, d in hols.items():
    if d == today:
        name = n; break
if name is None and today.weekday() >= 5:
    name = "Weekend"

# Run every day — Mark wants the map to refresh daily. On weekends/holidays the
# source data hasn't changed; `name` is kept for informative logging only and does
# NOT cause a skip.
#
# Idempotency guard: GitHub's `schedule` trigger is best-effort — runs are often
# delayed and sometimes dropped entirely. We therefore schedule the job twice
# (08:17 + a 09:17 catch-up). To keep the catch-up cheap, a SCHEDULED run exits
# immediately if today's data was already published (the tip of main is a
# delivery-map-bot "data: refresh <today>" commit). A manual workflow_dispatch
# NEVER skips, so Mark can always force a rebuild.
import subprocess

def already_published_today() -> bool:
    """True if the CURRENT tip of origin/main is today's bot data commit.

    Fetches rather than trusting the checked-out SHA: if the 08:17 run is still
    finishing when the 09:17 catch-up dispatches, the catch-up's checkout SHA can
    predate the bot's push. Fetching origin/main avoids a double publish.
    """
    try:
        subprocess.run(["git", "fetch", "--quiet", "--depth=1", "origin", "main"],
                       check=True, capture_output=True)
        out = subprocess.run(["git", "log", "-1", "--format=%an|%s", "FETCH_HEAD"],
                             capture_output=True, text=True, check=True).stdout.strip()
    except Exception:
        return False  # can't tell -> run rather than silently skip a day
    author, _, subject = out.partition("|")
    return author.strip() == "delivery-map-bot" and today.isoformat() in subject

event = os.environ.get("GITHUB_EVENT_NAME", "")
reason = name  # weekend / holiday name, informational only
skip = "false"
if event == "schedule" and already_published_today():
    skip = "true"
    reason = f"already published {today.isoformat()}"

out = os.environ.get("GITHUB_OUTPUT")
if out:
    with open(out, "a") as fh: fh.write(f"skip={skip}\n")
print(f"skip={skip}" + (f" ({reason})" if reason else "") + f" [event={event or 'n/a'}]")
